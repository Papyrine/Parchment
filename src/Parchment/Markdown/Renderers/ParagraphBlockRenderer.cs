class ParagraphBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, ParagraphBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, ParagraphBlock block)
    {
        ParagraphProperties? properties = null;
        var cls = MarkdownStyle.Resolve(block);
        if (cls != null)
        {
            properties = new()
            {
                ParagraphStyleId = new()
                {
                    Val = cls
                }
            };
        }

        renderer.WriteLeafInline(block);
        renderer.FlushParagraph(properties);
    }
}
