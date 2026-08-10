class RegisteredMarkdownTemplate(
    string name,
    Type modelType,
    byte[] styleSourceBytes,
    IReadOnlyList<PartScopeTree> nonBodyParts,
    IFluidTemplate parsedTemplate,
    ImagePolicies imagePolicies,
    IPageNumberResolver? pageNumbers,
    ExcelsiorTableMap excelsiorTables,
    IReadOnlyList<MarkdownExcelsiorTables.Placeholder> tablePlaceholders) :
    RegisteredTemplate(name, modelType)
{
    public override async Task Render(object model, Stream output, WordDocumentProperties? properties, Cancel cancel)
    {
        var context = new TemplateContext(model, SharedFluid.MarkdownOptions, allowModelMembers: true);
        await using var writer = new StringWriter();

        // Scoped to the liquid render alone: html reaching this flow is parked rather than written,
        // and only what runs inside the render can park it. The map is taken off the scope here and
        // passed explicitly from now on, so the ambient state ends with the render that owns it.
        IReadOnlyDictionary<string, string> htmlBlocks;
        using (var scope = MarkdownHtmlBlocks.BeginScope())
        {
            try
            {
                // MarkdownEncoder rather than NullEncoder: a bound value is data, so it is escaped
                // into the markdown being assembled rather than allowed to become part of its
                // syntax.
                await parsedTemplate.RenderAsync(writer, MarkdownEncoder.Default, context);
            }
            catch (TokenNotRenderableException exception)
            {
                // A Fluid value converter cannot see the template name, so it is attached here.
                throw new ParchmentRenderException(Name, exception.Message);
            }

            htmlBlocks = scope.Pending;
        }

        var markdownText = writer.ToString();
        cancel.ThrowIfCancellationRequested();

        using var stream = DocxCloner.ToWritableStream(styleSourceBytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var body = mainPart.Document!.Body
                ?? throw new ParchmentRenderException(Name, "Document has no body");

            var sectPr = body.Elements<SectionProperties>().LastOrDefault()
                         ?? body.Descendants<SectionProperties>().LastOrDefault();
            sectPr?.Remove();
            body.RemoveAllChildren();

            cancel.ThrowIfCancellationRequested();
            var numberingState = new WordNumberingState(mainPart);
            var elements = MarkdownRendering.Render(markdownText, mainPart, numberingState, imagePolicies, headingOffset: 0);
            foreach (var element in elements)
            {
                body.AppendChild(element);
            }

            // Before anything measures or numbers the document: a table is taller than the marker
            // paragraph it replaces, so every page number after it moves.
            MarkdownExcelsiorTables.Apply(body, tablePlaceholders, excelsiorTables, model, mainPart, Name);

            // Same reason, and the converted html is likewise taller than the marker it replaces.
            MarkdownHtmlBlocks.Apply(body, htmlBlocks, mainPart, imagePolicies, Name);

            if (sectPr != null)
            {
                body.AppendChild(sectPr);
            }

            // Before anything measures the document: a table of contents that grows from a
            // placeholder to its entries afterwards would move every heading below it.
            if (pageNumbers != null)
            {
                WordFieldResolution.WriteTableOfContents(body);
            }

            await RenderNonBodyParts(doc, mainPart, model, numberingState, cancel);

            // Stamp compatibilityMode=15 so Word opens the output normally instead of in
            // "Compatibility Mode" (a docx with no compat block is treated as Word 2007 / mode 12).
            SettingsCompatibility.Apply(mainPart);

            if (properties != null)
            {
                DocumentPropertiesWriter.Apply(doc, properties);
            }

            doc.Save();
        }

        // The document is complete, so it can be measured: whatever can paginate it reports where
        // each bookmark landed and the page numbers are written in, leaving Word nothing to ask
        // about when the file opens.
        if (pageNumbers != null)
        {
            await WordFieldResolution.Resolve(stream, pageNumbers, cancel);
        }

        // A part added during this render — settings here, numbering or images elsewhere — is
        // stamped with the wall clock, so byte-identical output held only when two renders landed
        // in the same 2-second zip quantum. See ZipTimestamps.
        ZipTimestamps.Pin(stream);
        stream.Position = 0;
        await stream.CopyToAsync(output, cancel);
    }

    // Headers, footers and notes come from the style source with their tokens intact, so they bind
    // here against the same model. The maps the docx flow threads through are all docx-only token
    // kinds, so they are empty: a header binds substitutions, loops and conditionals, nothing more.
    async Task RenderNonBodyParts(
        WordprocessingDocument doc,
        MainDocumentPart mainPart,
        object model,
        WordNumberingState numberingState,
        Cancel cancel)
    {
        if (nonBodyParts.Count == 0)
        {
            return;
        }

        var partRoots = new Dictionary<string, OpenXmlCompositeElement>(StringComparer.Ordinal);
        foreach (var (uri, root) in DocxCloner.EnumerateParts(doc))
        {
            partRoots[uri] = root;
        }

        var context = new TemplateContext(model, SharedFluid.Options, allowModelMembers: true);
        var styles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));

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
                numberingState,
                styles,
                imagePolicies);
            await runner.RunAsync(part.Nodes);
            runner.ApplyStructural();

            Anchors.StripAll(root);
            // A header or footer is a docx part, so its substitutions go through the text-based
            // path and its line breaks are materialized the same way the docx flow's are.
            LineBreaks.Apply(root);
        }
    }
}
