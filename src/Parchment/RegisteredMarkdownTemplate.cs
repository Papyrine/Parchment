class RegisteredMarkdownTemplate(
    string name,
    Type modelType,
    byte[] styleSourceBytes,
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
}
