[RequiresUnreferencedCode(Unsupported.Trimming)]
[RequiresDynamicCode(Unsupported.DynamicCode)]
class RegisteredDocxTemplate(
    string name,
    Type modelType,
    byte[] canonicalBytes,
    IReadOnlyList<PartScopeTree> parts,
    ExcelsiorTableMap excelsiorTables,
    FormatMap formats,
    StringListMap stringLists,
    EditableMap editables,
    ImagePolicies imagePolicies,
    int maxBookmarkId,
    NoteReferences noteReferences,
    TemplateIds templateIds,
    int maxEditableId) :
    RegisteredTemplate(name, modelType)
{
    // The parts a render has work in. A part outside this set carries no anchor bookmark, so
    // nothing in it is substituted, removed or produced, and it is never read: it stays as the
    // bytes cloned from the registration snapshot. See "Render reads only the parts that carry
    // tokens" in contributing.md for what registration does to make that hold.
    readonly HashSet<string> tokenParts = parts
        .Select(_ => _.PartUri)
        .ToHashSet(StringComparer.Ordinal);

    internal byte[] Snapshot => canonicalBytes;

    // The token parts, parsed on the first render and cloned by every render after it instead of being
    // parsed from xml again. Shared by concurrent renders and never mutated - see ParsedParts.
    ParsedParts? parsedParts;
    object? parsedPartsLock;

    public override async Task Render(object model, Stream output, WordDocumentProperties? properties, Cancel cancel)
    {
        cancel.ThrowIfCancellationRequested();

        var context = SharedFluid.CreateContext(model, SharedFluid.Options);
        var parsed = LazyInitializer.EnsureInitialized(
            ref parsedParts,
            ref parsedPartsLock,
            () => ParsedParts.Read(canonicalBytes, tokenParts));
        using var stream = DocxCloner.ToWritableStream(parsed.Bytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var mainPartUri = mainPart.Uri.ToString();
            // Materialize the part roots once, keyed by uri. The previous code re-enumerated all
            // parts (re-allocating each part's Uri.ToString()) to locate every part's root, and
            // enumerated a second time for the strip pass — O(parts²) plus redundant string allocs.
            // Only the token-bearing parts: reading a root parses the part and the save then writes
            // it back out, which for a token-free header, footer or notes part is all overhead.
            var partRoots = new Dictionary<string, (OpenXmlPart part, OpenXmlCompositeElement root)>(StringComparer.Ordinal);
            foreach (var (uri, part, root) in DocxCloner.EnumerateParts(doc, tokenParts))
            {
                // The root arrives empty: its content is a copy of the tree kept from the first
                // render, which costs less than parsing the same xml again.
                parsed.Fill(uri, root);
                partRoots[uri] = (part, root);
            }

            var numberingState = new WordNumberingState(mainPart);
            var editableState = new EditableState(maxEditableId);
            var bookmarks = new BookmarkIds(maxBookmarkId);
            var cloneIds = CloneIds.For(templateIds);
            // Cache the StyleSet per render — Excelsior / Format / OpenXml / Mutate tokens all
            // need it, but it changes only across registrations, not within a render.
            var styles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));

            try
            {
                foreach (var part in parts)
                {
                    cancel.ThrowIfCancellationRequested();
                    if (!partRoots.TryGetValue(part.PartUri, out var target))
                    {
                        continue;
                    }

                    await RenderPartAsync(mainPart, part, target.part, target.root, part.PartUri == mainPartUri, context, model, numberingState, editableState, bookmarks, cloneIds, styles, cancel).ConfigureAwait(false);
                }

                foreach (var (_, root) in partRoots.Values)
                {
                    cancel.ThrowIfCancellationRequested();
                    Anchors.StripAll(root);
                    // Once every substitution in the part has run, so splitting a w:t can no longer
                    // disturb the offsets the substitutions were driven by.
                    LineBreaks.Apply(root);
                    // A loop copies a comment's anchors with its body, and a comment has one range.
                    if (templateIds.HoldsComments)
                    {
                        CommentAnchors.KeepFirst(root);
                    }

                    // A removed branch, an empty loop or a structural token can take a cell's only
                    // paragraph with it, or leave a nested table as the cell's last block.
                    TableCells.EnsureFinalParagraph(root);
                }

                // After every part has rendered and been cleaned up, the notes parts included: a loop
                // repeats a footnote reference and an unchosen branch removes one, and the copy of a
                // note made here has to be of its final content. A body with no scope tree was never
                // touched, so its references still pair with their notes.
                if (!noteReferences.IsEmpty &&
                    partRoots.TryGetValue(mainPartUri, out var body))
                {
                    noteReferences.Apply(mainPart, body.root);
                }
            }
            catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
            {
                // The last resort: a failure with a token or a tag to name was wrapped where it
                // happened, and anything else inside a part by RenderPartAsync. Document properties
                // are applied outside it - they are the caller's argument, not the template's.
                throw RenderFailure.Wrap(exception, Name);
            }

            // compatibilityMode=15 is baked into the registration snapshot (see
            // TemplateStore.RegisterDocxTemplate), so no per-render stamp is needed here.
            if (properties != null)
            {
                DocumentPropertiesWriter.Apply(doc, properties, Name);
            }

            cancel.ThrowIfCancellationRequested();
            doc.Save();
        }

        // A part added during this render — a numbering part for a list, an image, editable-field
        // seeding — is stamped with the wall clock, so byte-identical output held only when two
        // renders landed in the same 2-second zip quantum. See ZipTimestamps.
        ZipTimestamps.Pin(stream);
        stream.Position = 0;
        await stream.CopyToAsync(output, cancel).ConfigureAwait(false);
    }

    async Task RenderPartAsync(MainDocumentPart mainPart, PartScopeTree part, OpenXmlPart owner, OpenXmlCompositeElement root, bool isBody, TemplateContext context, object model, WordNumberingState numberingState, EditableState editableState, BookmarkIds bookmarks, CloneIds? cloneIds, Lazy<StyleSet> styles, Cancel cancel)
    {
        // Editable fields dispatch only in the document body. Word does not reliably honor
        // editable-range exceptions in headers/footers/notes, so the same token there renders
        // as plain read-only text — deliberate (e.g. an editable PO number in the body can be
        // mirrored read-only in the footer).
        var map = Anchors.BuildMap(root);
        // A link or an image produced in a header, footer or note has to hang off that part.
        var relationships = new PartRelationships(mainPart, owner);
        var runner = new ScopeTreeRunner(
            Name,
            part.PartUri,
            map,
            context,
            mainPart,
            relationships,
            model,
            excelsiorTables,
            formats,
            stringLists,
            isBody ? editables : EditableMap.Empty,
            editableState,
            bookmarks,
            cloneIds,
            numberingState,
            styles,
            imagePolicies,
            cancel);
        try
        {
            await runner.RunAsync(part.Nodes).ConfigureAwait(false);
            runner.ApplyStructural();
            runner.SettleSectionBreaks();
            runner.SeparateTables();
            runner.ApplySectionBreaks(root);
            relationships.Complete();
        }
        catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
        {
            throw RenderFailure.Wrap(exception, Name, part.PartUri);
        }
    }
}
