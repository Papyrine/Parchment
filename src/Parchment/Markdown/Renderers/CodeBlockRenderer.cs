class CodeBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, CodeBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, CodeBlock block)
    {
        var indent = renderer.CurrentIndent;
        var count = block.Lines.Count;
        if (count == 0)
        {
            return;
        }

        var styleId = MarkdownStyle.ResolveCodeBlock(block) ?? "Code";

        foreach (var line in block.Lines.Lines.Take(count))
        {
            var properties = new ParagraphProperties(
                new ParagraphStyleId
                {
                    Val = styleId
                });
            if (indent > 0)
            {
                properties.Append(
                    new Indentation
                    {
                        Left = indent.ToString(CultureInfo.InvariantCulture)
                    });
            }

            var paragraph = new Paragraph(
                properties,
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
