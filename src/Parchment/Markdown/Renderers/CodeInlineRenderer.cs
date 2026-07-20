class CodeInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, CodeInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, CodeInline inline)
    {
        var properties = new RunProperties(
            new RunFonts
            {
                Ascii = "Consolas",
                HighAnsi = "Consolas"
            });

        // Assigned through the typed property so rStyle leads the rPr sequence.
        if (MarkdownStyle.Resolve(inline) is { } styleId)
        {
            properties.RunStyle = new()
            {
                Val = styleId
            };
        }

        var run = new Run(
            properties,
            new Text(XmlCharSanitizer.Strip(inline.ContentSpan).ToString())
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        renderer.AddRun(run);
    }
}
