class RegisteredMarkdownTemplate(
    string name,
    Type modelType,
    byte[] styleSourceBytes,
    IReadOnlyList<PartScopeTree> nonBodyParts,
    IFluidTemplate parsedTemplate,
    ImagePolicies imagePolicies) :
    RegisteredTemplate(name, modelType)
{
    public override async Task Render(object model, Stream output, WordDocumentProperties? properties, Cancel cancel)
    {
        var context = new TemplateContext(model, SharedFluid.MarkdownOptions, allowModelMembers: true);
        await using var writer = new StringWriter();
        try
        {
            await parsedTemplate.RenderAsync(writer, NullEncoder.Default, context);
        }
        catch (TokenNotRenderableException exception)
        {
            // A Fluid value converter cannot see the template name, so it is attached here.
            throw new ParchmentRenderException(Name, exception.Message);
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

            if (sectPr != null)
            {
                body.AppendChild(sectPr);
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
        }
    }
}
