class QuoteBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, QuoteBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, QuoteBlock quoteBlock)
    {
        // `{.MyQuote}` on its own line above the quote binds to the QuoteBlock and covers every
        // line of it; written at the end of a quoted line it binds to that line's paragraph and
        // covers only that one. The narrower of the two wins, matching how a list and its items
        // resolve.
        var blockStyle = MarkdownStyle.Resolve(quoteBlock);

        foreach (var child in quoteBlock)
        {
            if (child is LeafBlock leaf and
                not HtmlBlock and
                not CodeBlock and
                not ThematicBreakBlock)
            {
                var properties = new ParagraphProperties
                {
                    ParagraphStyleId = new()
                    {
                        Val = MarkdownStyle.Resolve(leaf) ?? blockStyle ?? "Quote"
                    }
                };
                renderer.WriteLeafInline(leaf);
                renderer.FlushParagraph(properties);
            }
            else
            {
                renderer.PushIndent(720);
                try
                {
                    renderer.Render(child);
                }
                finally
                {
                    renderer.PopIndent();
                }
            }
        }
    }
}
