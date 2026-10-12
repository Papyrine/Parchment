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
    int maxBookmarkId) :
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

    public override async Task Render(object model, Stream output, WordDocumentProperties? properties, Cancel cancel)
    {
        cancel.ThrowIfCancellationRequested();

        var context = new TemplateContext(model, SharedFluid.Options, allowModelMembers: true);
        using var stream = DocxCloner.ToWritableStream(canonicalBytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var mainPartUri = mainPart.Uri.ToString();
            // Materialize the part roots once, keyed by uri. The previous code re-enumerated all
            // parts (re-allocating each part's Uri.ToString()) to locate every part's root, and
            // enumerated a second time for the strip pass — O(parts²) plus redundant string allocs.
            // Only the token-bearing parts: reading a root parses the part and the save then writes
            // it back out, which for a token-free header, footer or notes part is all overhead.
            var partRoots = new Dictionary<string, OpenXmlCompositeElement>(StringComparer.Ordinal);
            foreach (var (uri, root) in DocxCloner.EnumerateParts(doc, tokenParts))
            {
                partRoots[uri] = root;
            }

            var numberingState = new WordNumberingState(mainPart);
            var editableState = new EditableState();
            var bookmarks = new BookmarkIds(maxBookmarkId);
            // Cache the StyleSet per render — Excelsior / Format / OpenXml / Mutate tokens all
            // need it, but it changes only across registrations, not within a render.
            var styles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));

            foreach (var part in parts)
            {
                cancel.ThrowIfCancellationRequested();
                if (!partRoots.TryGetValue(part.PartUri, out var root))
                {
                    continue;
                }

                await RenderPartAsync(mainPart, part, root, part.PartUri == mainPartUri, context, model, numberingState, editableState, bookmarks, styles);
            }

            foreach (var root in partRoots.Values)
            {
                Anchors.StripAll(root);
                // Once every substitution in the part has run, so splitting a w:t can no longer
                // disturb the offsets the substitutions were driven by.
                LineBreaks.Apply(root);
                // A removed branch, an empty loop or a structural token can take a cell's only
                // paragraph with it, or leave a nested table as the cell's last block.
                TableCells.EnsureFinalParagraph(root);
            }

            // compatibilityMode=15 is baked into the registration snapshot (see
            // TemplateStore.RegisterDocxTemplate), so no per-render stamp is needed here.
            if (properties != null)
            {
                DocumentPropertiesWriter.Apply(doc, properties);
            }

            doc.Save();
        }

        // A part added during this render — a numbering part for a list, an image, editable-field
        // seeding — is stamped with the wall clock, so byte-identical output held only when two
        // renders landed in the same 2-second zip quantum. See ZipTimestamps.
        ZipTimestamps.Pin(stream);
        stream.Position = 0;
        await stream.CopyToAsync(output, cancel);
    }

    async Task RenderPartAsync(MainDocumentPart mainPart, PartScopeTree part, OpenXmlCompositeElement root, bool isBody, TemplateContext context, object model, WordNumberingState numberingState, EditableState editableState, BookmarkIds bookmarks, Lazy<StyleSet> styles)
    {
        // Editable fields dispatch only in the document body. Word does not reliably honor
        // editable-range exceptions in headers/footers/notes, so the same token there renders
        // as plain read-only text — deliberate (e.g. an editable PO number in the body can be
        // mirrored read-only in the footer).
        var map = Anchors.BuildMap(root);
        var runner = new ScopeTreeRunner(
            Name,
            part.PartUri,
            map,
            context,
            mainPart,
            model,
            excelsiorTables,
            formats,
            stringLists,
            isBody ? editables : EditableMap.Empty,
            editableState,
            bookmarks,
            numberingState,
            styles,
            imagePolicies);
        await runner.RunAsync(part.Nodes);
        runner.ApplyStructural();
        runner.SettleSectionBreaks();
        runner.SeparateProducedTables();
    }
}
