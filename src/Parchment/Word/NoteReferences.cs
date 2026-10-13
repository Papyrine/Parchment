/// <summary>
/// Keeps footnotes and endnotes one to one with the references a render leaves in the body.
/// </summary>
/// <remarks>
/// <para>
/// A note and its reference live in different parts and are tied by id alone. Word refuses to open
/// a document in which two references name one note, and <c>OpenXmlValidator</c> does not report
/// it, because the schema has no such rule.
/// </para>
/// <para>
/// A render breaks the pairing from the body's side only. A <c>{% for %}</c> clones the reference
/// with every iteration, and an empty loop or a branch that was not chosen removes it and leaves
/// its note behind. Those are many sites, nested to any depth, so the repair is one pass over the
/// body once every part has been rendered: each reference after the first to a note gets a copy of
/// that note under a fresh id, and a note the template referred to that no reference names any
/// more is removed. Running after the notes part has been rendered is what lets a copy carry the
/// note's substituted text instead of its tokens and anchors.
/// </para>
/// <para>
/// The ids the template's body refers to are read at registration. A template with none — nearly
/// every template — skips the pass, and one with some walks its body but reads a notes part only
/// when a count actually moved, so a token-free notes part stays the bytes of the snapshot.
/// </para>
/// </remarks>
sealed class NoteReferences
{
    public static NoteReferences None { get; } = new([], []);

    readonly HashSet<long> footnotes;
    readonly HashSet<long> endnotes;

    NoteReferences(HashSet<long> footnotes, HashSet<long> endnotes)
    {
        this.footnotes = footnotes;
        this.endnotes = endnotes;
    }

    public bool IsEmpty =>
        footnotes.Count == 0 &&
        endnotes.Count == 0;

    /// <summary>
    /// The notes <paramref name="body"/> refers to, as the template wrote it.
    /// </summary>
    public static NoteReferences Read(MainDocumentPart mainPart, OpenXmlCompositeElement body)
    {
        // A reference can only name a note in a part that exists, so a template with neither part
        // is not walked.
        if (mainPart.FootnotesPart == null &&
            mainPart.EndnotesPart == null)
        {
            return None;
        }

        HashSet<long>? footnotes = null;
        HashSet<long>? endnotes = null;
        foreach (var element in body.Descendants())
        {
            if (element is FootnoteReference {Id.Value: var footnote})
            {
                (footnotes ??= []).Add(footnote);
            }
            else if (element is EndnoteReference {Id.Value: var endnote})
            {
                (endnotes ??= []).Add(endnote);
            }
        }

        if (footnotes == null &&
            endnotes == null)
        {
            return None;
        }

        return new(footnotes ?? [], endnotes ?? []);
    }

    /// <summary>
    /// Gives every reference in the rendered <paramref name="body"/> a note of its own, and removes
    /// the notes whose references the render took out.
    /// </summary>
    public void Apply(MainDocumentPart mainPart, OpenXmlCompositeElement body)
    {
        List<FootnoteEndnoteReferenceType>? footnoteReferences = null;
        List<FootnoteEndnoteReferenceType>? endnoteReferences = null;
        foreach (var element in body.Descendants())
        {
            if (element is FootnoteReference footnote)
            {
                (footnoteReferences ??= []).Add(footnote);
            }
            else if (element is EndnoteReference endnote)
            {
                (endnoteReferences ??= []).Add(endnote);
            }
        }

        Pair(footnotes, footnoteReferences, mainPart, static _ => _.FootnotesPart?.Footnotes);
        Pair(endnotes, endnoteReferences, mainPart, static _ => _.EndnotesPart?.Endnotes);
    }

    static void Pair(
        HashSet<long> template,
        List<FootnoteEndnoteReferenceType>? references,
        MainDocumentPart mainPart,
        Func<MainDocumentPart, OpenXmlCompositeElement?> notesRoot)
    {
        // In reading order, so the first reference to a note is the one that keeps it and the
        // copies are numbered the way the document reads.
        HashSet<long>? seen = null;
        List<FootnoteEndnoteReferenceType>? repeats = null;
        if (references != null)
        {
            foreach (var reference in references)
            {
                if (reference.Id?.Value is not { } id)
                {
                    continue;
                }

                if (!(seen ??= []).Add(id))
                {
                    (repeats ??= []).Add(reference);
                }
            }
        }

        List<long>? orphaned = null;
        foreach (var id in template)
        {
            if (seen == null ||
                !seen.Contains(id))
            {
                (orphaned ??= []).Add(id);
            }
        }

        if (repeats == null &&
            orphaned == null)
        {
            return;
        }

        // Only now: reading the root parses the part, and the save then writes it back out.
        var root = notesRoot(mainPart);
        if (root == null)
        {
            return;
        }

        var notes = new Dictionary<long, FootnoteEndnoteType>();
        var last = 0L;
        foreach (var note in root.Elements<FootnoteEndnoteType>())
        {
            if (note.Id?.Value is not { } id)
            {
                continue;
            }

            notes.TryAdd(id, note);
            last = Math.Max(last, id);
        }

        if (repeats != null)
        {
            // Each copy goes after the previous copy of the same note, so one note's copies sit
            // together in the part, in the order their references are read.
            var tails = new Dictionary<long, FootnoteEndnoteType>();
            foreach (var reference in repeats)
            {
                var id = reference.Id!.Value;
                if (!notes.TryGetValue(id, out var note))
                {
                    continue;
                }

                var copy = (FootnoteEndnoteType)note.CloneNode(true);
                copy.Id = ++last;
                tails.GetValueOrDefault(id, note).InsertAfterSelf(copy);
                tails[id] = copy;
                reference.Id = last;
            }
        }

        if (orphaned != null)
        {
            foreach (var id in orphaned)
            {
                // The separators are notes too, and are referred to from the settings, not the body.
                if (notes.TryGetValue(id, out var note) &&
                    (note.Type == null ||
                     note.Type.Value == FootnoteEndnoteValues.Normal))
                {
                    note.Remove();
                }
            }
        }
    }
}
