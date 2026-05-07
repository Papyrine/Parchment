class CodeBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, CodeBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, CodeBlock block)
    {
        foreach (var line in block.Lines.Lines.Take(block.Lines.Count))
        {
            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId
                    {
                        Val = "Code"
                    }),
                new Run(
                    new RunProperties(
                        new RunFonts
                        {
                            Ascii = "Consolas",
                            HighAnsi = "Consolas"
                        }),
                    new Text(XmlCharSanitizer.Strip(line.Slice.AsSpan()).ToString())
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
            renderer.AddBlock(paragraph);
        }
    }
}
