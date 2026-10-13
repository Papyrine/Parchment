[RequiresUnreferencedCode(Unsupported.Trimming)]
[RequiresDynamicCode(Unsupported.DynamicCode)]
class RegisteredMarkdownTemplate(
    string name,
    Type modelType,
    byte[] styleSourceBytes,
    IReadOnlyList<PartScopeTree> nonBodyParts,
    IFluidTemplate parsedTemplate,
    ImagePolicies imagePolicies,
    IPageNumberResolver? pageNumbers,
    ExcelsiorTableMap excelsiorTables,
    IReadOnlyList<MarkdownExcelsiorTables.Placeholder> tablePlaceholders,
    int maxBookmarkId,
    TemplateIds templateIds,
    bool compatibilityBaked) :
    RegisteredTemplate(name, modelType)
{
    // As in the docx flow: a non-body part outside this set carries no anchor bookmark and is
    // never read, so it stays as the bytes cloned from the style source.
    readonly HashSet<string> tokenParts = nonBodyParts
        .Select(_ => _.PartUri)
        .ToHashSet(StringComparer.Ordinal);

    internal byte[] Snapshot => styleSourceBytes;

    public override async Task Render(object model, Stream output, WordDocumentProperties? properties, Cancel cancel)
    {
        var context = new TemplateContext(model, SharedFluid.MarkdownOptions, allowModelMembers: true);
        using var writer = new StringWriter();

        // Scoped to the liquid render alone: a value with no markdown form is parked rather than
        // written, and only what runs inside the render can park it. The map is taken off the scope here and
        // passed explicitly from now on, so the ambient state ends with the render that owns it.
        IReadOnlyDictionary<string, TokenValue> tokenBlocks;
        using (var scope = MarkdownTokenBlocks.BeginScope())
        {
            try
            {
                // MarkdownEncoder rather than NullEncoder: a bound value is data, so it is escaped
                // into the markdown being assembled rather than allowed to become part of its
                // syntax.
                await parsedTemplate.RenderAsync(writer, MarkdownEncoder.Default, context).ConfigureAwait(false);
            }
            catch (TokenNotRenderableException exception)
            {
                // A Fluid value converter cannot see the template name, so it is attached here.
                throw new ParchmentRenderException(Name, exception.Message);
            }

            tokenBlocks = scope.Pending;
        }

        var markdownText = writer.ToString();
        cancel.ThrowIfCancellationRequested();

        using var stream = DocxCloner.ToWritableStream(styleSourceBytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var body = mainPart.Document!.Body
                ?? throw new ParchmentRenderException(Name, "Document has no body");

            // The snapshot's body was emptied at registration and holds nothing but the section to
            // preserve - see TemplateStore.BakeStyleSource. It is detached from this render's own
            // copy, because SectionBreaks rewrites it in place.
            var sectPr = body.GetFirstChild<SectionProperties>();
            sectPr?.Remove();

            cancel.ThrowIfCancellationRequested();
            var numberingState = new WordNumberingState(mainPart);
            // Lazy because only an OpenXmlToken asks for it, and most renders have none.
            var bodyStyles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));
            // One counter for the whole render, seeded above what the headers, footers and notes
            // already carry: the body is rebuilt from nothing, but those parts arrive as authored.
            var bookmarks = new BookmarkIds(maxBookmarkId);
            IReadOnlyList<OpenXmlElement> elements;
            bool mayHoldSections;
            try
            {
                elements = MarkdownRendering.Render(markdownText, mainPart, numberingState, imagePolicies, bookmarks, headingOffset: 0, out mayHoldSections);
            }
            catch (SectionMarkerException exception)
            {
                // The markdown walk cannot see the template name, so it is attached here.
                throw new ParchmentRenderException(Name, exception.Message);
            }
            catch (TableColumnLimitException exception)
            {
                // The table renderer cannot see the template name either, so it is attached here.
                throw new ParchmentRenderException(Name, exception.Message);
            }

            foreach (var element in elements)
            {
                body.AppendChild(element);
            }

            // Before anything measures or numbers the document: a table is taller than the marker
            // paragraph it replaces, so every page number after it moves.
            MarkdownExcelsiorTables.Apply(body, tablePlaceholders, excelsiorTables, model, mainPart, bookmarks, Name);

            // Same reason, and what a parked token produces is likewise taller than the marker it
            // replaces.
            MarkdownTokenBlocks.Apply(body, tokenBlocks, mainPart, numberingState, bodyStyles, imagePolicies, bookmarks, Name);

            // A pipe-table cell whose value rendered to a table ends with that table, and Word
            // wants a paragraph after it. Here rather than later because the added paragraph has
            // height, like the swaps above.
            TableCells.EnsureFinalParagraph(body);

            // Before the trailing section is attached, because the last [SECTION] in the template
            // is what that section describes — see SectionBreaks. Skipped when nothing in the body can
            // be a marker: the markdown walk wrote none and converted no html, and neither swap above
            // ran, those being where caller code - an OpenXmlToken, an Excelsior Configure hook - can
            // put in a paragraph this render did not build.
            if (mayHoldSections ||
                tokenBlocks.Count > 0 ||
                tablePlaceholders.Count > 0)
            {
                SectionBreaks.Apply(body, sectPr, Name);
            }

            if (sectPr != null)
            {
                body.AppendChild(sectPr);
            }

            // Before anything measures the document: a table of contents that grows from a
            // placeholder to its entries afterwards would move every heading below it.
            if (pageNumbers != null)
            {
                WordFieldResolution.WriteTableOfContents(body, bookmarks);
            }

            await RenderNonBodyParts(doc, mainPart, model, numberingState, bookmarks, cancel).ConfigureAwait(false);

            // Stamp compatibilityMode=15 so Word opens the output normally instead of in
            // "Compatibility Mode" (a docx with no compat block is treated as Word 2007 / mode 12).
            // A style source with a settings part of its own had it stamped at registration.
            if (!compatibilityBaked)
            {
                SettingsCompatibility.Apply(mainPart);
            }

            if (properties != null)
            {
                DocumentPropertiesWriter.Apply(doc, properties, Name);
            }

            doc.Save();
        }

        // The document is complete, so it can be measured: whatever can paginate it reports where
        // each bookmark landed and the page numbers are written in, leaving Word nothing to ask
        // about when the file opens.
        if (pageNumbers != null)
        {
            await WordFieldResolution.Resolve(stream, pageNumbers, cancel).ConfigureAwait(false);
        }

        // A part added during this render — settings here, numbering or images elsewhere — is
        // stamped with the wall clock, so byte-identical output held only when two renders landed
        // in the same 2-second zip quantum. See ZipTimestamps.
        ZipTimestamps.Pin(stream);
        stream.Position = 0;
        await stream.CopyToAsync(output, cancel).ConfigureAwait(false);
    }

    // Headers, footers and notes come from the style source with their tokens intact, so they bind
    // here against the same model. The maps the docx flow threads through are all docx-only token
    // kinds, so they are empty: a header binds substitutions, loops and conditionals, nothing more.
    async Task RenderNonBodyParts(
        WordprocessingDocument doc,
        MainDocumentPart mainPart,
        object model,
        WordNumberingState numberingState,
        BookmarkIds bookmarks,
        Cancel cancel)
    {
        if (nonBodyParts.Count == 0)
        {
            return;
        }

        var partRoots = new Dictionary<string, OpenXmlCompositeElement>(StringComparer.Ordinal);
        foreach (var (uri, root) in DocxCloner.EnumerateParts(doc, tokenParts))
        {
            partRoots[uri] = root;
        }

        var context = new TemplateContext(model, SharedFluid.Options, allowModelMembers: true);
        var styles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));
        // One for all of them, as with the bookmarks: the ids are unique across the document.
        var cloneIds = CloneIds.For(templateIds);

        foreach (var part in nonBodyParts)
        {
            cancel.ThrowIfCancellationRequested();
            if (!partRoots.TryGetValue(part.PartUri, out var root))
            {
                continue;
            }

            var runner = new ScopeTreeRunner(
                Name,
                part.PartUri,
                Anchors.BuildMap(root),
                context,
                mainPart,
                model,
                ExcelsiorTableMap.Empty,
                FormatMap.Empty,
                StringListMap.Empty,
                EditableMap.Empty,
                new(),
                bookmarks,
                cloneIds,
                numberingState,
                styles,
                imagePolicies);
            await runner.RunAsync(part.Nodes).ConfigureAwait(false);
            runner.ApplyStructural();
            runner.SeparateTables();
            // None of these parts is the body, so this only ever rejects a marker - the body's own
            // are finished by SectionBreaks.Apply in Render.
            runner.ApplySectionBreaks(root);

            Anchors.StripAll(root);
            // A header or footer is a docx part, so its substitutions go through the text-based
            // path and its line breaks are materialized the same way the docx flow's are.
            LineBreaks.Apply(root);
            // A loop copies a comment's anchors with its body, here as in the docx flow.
            if (templateIds.HoldsComments)
            {
                CommentAnchors.KeepFirst(root);
            }

            // And a header or footer table cell can lose its only paragraph the same way.
            TableCells.EnsureFinalParagraph(root);
        }
    }
}
