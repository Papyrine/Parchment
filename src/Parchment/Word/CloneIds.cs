using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using StringValue = DocumentFormat.OpenXml.StringValue;

/// <summary>
/// Per-render allocator for the ids a copy of template content cannot share with what it was
/// copied from: a drawing's <c>wp:docPr</c> id, a content control's <c>w:id</c>, and the
/// <c>w:id</c> of a tracked change.
/// </summary>
/// <remarks>
/// <para>
/// A render copies template content in two places. A <c>{% for %}</c> clones its body once per
/// item, and <see cref="ParagraphSplicer"/> clones a paragraph to cut it in two around a token.
/// Each copy arrives with the ids of its source, and all three kinds have to be unique. None of
/// them is referred to by anything, so the copy is given new ones: the first copy of a loop body
/// keeps the template's, every later one is renumbered as it is made, and the after half of a
/// cut paragraph is renumbered where it still shares an id with the before half. As with
/// <see cref="BookmarkIds"/>, only the copy is walked, never the document it lands in.
/// </para>
/// <para>
/// The counters start from what <see cref="TemplateIds"/> read at registration and run in
/// processing order, so the same template and model give the same ids. Drawings and tracked
/// changes count up from the template's highest. Content controls count down from its lowest, or
/// from zero: an editable field numbers the controls it creates upwards from one
/// (<see cref="EditableState"/>), and the two must not meet. Word draws these ids from the whole
/// signed range, so a negative one is ordinary.
/// </para>
/// <para>
/// Not covered: a marker that is half of a pair. A template bookmark, a moved or customXml range
/// and an editable range each tie a start to an end by id, and a loop still copies those as they
/// are. A comment's anchors are not renumbered either, since the id is what names the comment —
/// see <see cref="CommentAnchors"/>.
/// </para>
/// </remarks>
sealed class CloneIds
{
    uint lastDrawing;
    int lastControl;
    int lastRevision;

    CloneIds(TemplateIds template)
    {
        lastDrawing = template.MaxDrawing;
        lastControl = template.MinControl;
        lastRevision = template.MaxRevision;
    }

    /// <summary>
    /// The allocator for one render, or null for a template that holds none of these ids — most
    /// of them — so that a render of one never walks a copy to look.
    /// </summary>
    public static CloneIds? For(TemplateIds template)
    {
        if (template.HoldsIds)
        {
            return new(template);
        }

        return null;
    }

    /// <summary>
    /// True when a second copy of <paramref name="elements"/> would repeat an id of the first.
    /// </summary>
    public static bool Holds(IReadOnlyList<OpenXmlElement> elements)
    {
        foreach (var element in elements)
        {
            foreach (var candidate in Tree(element))
            {
                if (IdOf(candidate) != null)
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Give everything in <paramref name="copy"/> a fresh id.
    /// </summary>
    public void Renumber(OpenXmlElement copy)
    {
        foreach (var element in Tree(copy))
        {
            Assign(element);
        }
    }

    /// <summary>
    /// Give <paramref name="copy"/> a fresh id wherever it carries one <paramref name="kept"/>
    /// carries too. For two halves cut from the same paragraph, where content placed in one half
    /// only — a drawing — is not a repeat and keeps the id it was authored with.
    /// </summary>
    public void RenumberShared(IReadOnlyList<OpenXmlElement> kept, IReadOnlyList<OpenXmlElement> copy)
    {
        HashSet<(Type, string)>? held = null;
        foreach (var element in kept)
        {
            foreach (var candidate in Tree(element))
            {
                if (IdOf(candidate) is { } id)
                {
                    held ??= [];
                    held.Add((candidate.GetType(), id));
                }
            }
        }

        if (held == null)
        {
            return;
        }

        foreach (var element in copy)
        {
            foreach (var candidate in Tree(element))
            {
                if (IdOf(candidate) is { } id &&
                    held.Contains((candidate.GetType(), id)))
                {
                    Assign(candidate);
                }
            }
        }
    }

    void Assign(OpenXmlElement element)
    {
        if (element is DW.DocProperties drawing)
        {
            drawing.Id = ++lastDrawing;
            return;
        }

        if (element is SdtId control)
        {
            control.Val = --lastControl;
            return;
        }

        if (RevisionId(element) is { } revision)
        {
            revision.Value = (++lastRevision).ToString(CultureInfo.InvariantCulture);
        }
    }

    static string? IdOf(OpenXmlElement element) =>
        element switch
        {
            DW.DocProperties drawing => drawing.Id?.InnerText,
            SdtId control => control.Val?.InnerText,
            _ => RevisionId(element)?.Value
        };

    /// <summary>
    /// The id of a tracked change — an insertion, a deletion, either side of a move, a change to
    /// formatting or to a table — or null for any other element. The classes share no base, so
    /// they are listed. The start of a customXml range is a <see cref="TrackChangeType"/> as well
    /// and is turned away: its end carries the same id.
    /// </summary>
    internal static StringValue? RevisionId(OpenXmlElement element) =>
        element switch
        {
            RunTrackChangeType change => change.Id,
            CustomXmlInsRangeStart or CustomXmlDelRangeStart or CustomXmlMoveFromRangeStart or CustomXmlMoveToRangeStart => null,
            TrackChangeType change => change.Id,
            RunPropertiesChange change => change.Id,
            ParagraphMarkRunPropertiesChange change => change.Id,
            ParagraphPropertiesChange change => change.Id,
            SectionPropertiesChange change => change.Id,
            TablePropertiesChange change => change.Id,
            TablePropertyExceptionsChange change => change.Id,
            TableGridChange change => change.Id,
            TableRowPropertiesChange change => change.Id,
            TableCellPropertiesChange change => change.Id,
            CellMerge change => change.Id,
            NumberingChange change => change.Id,
            MathControlMoveType change => change.Id,
            InsertedMathControl change => change.Id,
            DeletedMathControl change => change.Id,
            _ => null
        };

    // The element itself and then its descendants: what a loop copies can be a body-level marker
    // or wrapper as well as a paragraph or a table.
    static IEnumerable<OpenXmlElement> Tree(OpenXmlElement element)
    {
        yield return element;
        foreach (var descendant in element.Descendants())
        {
            yield return descendant;
        }
    }
}
