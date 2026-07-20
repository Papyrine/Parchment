class HeadingBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HeadingBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, HeadingBlock block)
    {
        var level = Math.Clamp(block.Level + renderer.HeadingOffset, 1, 9);
        var styleId = ResolveStyle(block, level);
        var properties = new ParagraphProperties
        {
            ParagraphStyleId = new()
            {
                Val = styleId
            }
        };
        renderer.WriteLeafInline(block);
        renderer.FlushParagraph(properties);
    }

    static string ResolveStyle(HeadingBlock block, int level) =>
        MarkdownStyle.Resolve(block) ?? $"Heading{level}";
}