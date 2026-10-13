class CodeInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, CodeInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, CodeInline inline)
    {
        // rStyle leads the rPr sequence, so it goes in before the fonts.
        var properties = new RunProperties();
        if (MarkdownStyle.Resolve(inline) is { } styleId)
        {
            properties.Append(
                new RunStyle
                {
                    Val = styleId
                });
        }

        properties.Append(
            new RunFonts
            {
                Ascii = "Consolas",
                HighAnsi = "Consolas"
            });

        var run = new Run(
            properties,
            new Text(XmlChars.Strip(inline.ContentSpan).ToString())
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        renderer.AddRun(run);
    }
}
