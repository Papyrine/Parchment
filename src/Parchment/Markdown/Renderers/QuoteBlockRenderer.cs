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
        var depth = ResolveDepth(quoteBlock);

        foreach (var child in quoteBlock)
        {
            if (child is HeadingBlock heading)
            {
                // The heading keeps its own style, so it is stated at the quote's indent directly.
                HeadingBlockRenderer.Write(
                    renderer,
                    heading,
                    new()
                    {
                        Indentation = new()
                        {
                            Left = Indent(depth)
                        }
                    });
                continue;
            }

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
                if (depth > 0)
                {
                    // The style indents every quote by the same amount, so a quote inside a quote
                    // sat level with the one around it. Each level in states its own indent.
                    properties.Indentation = new()
                    {
                        Left = Indent(depth)
                    };
                }

                renderer.WriteLeafInline(leaf);
                renderer.FlushParagraph(properties);
            }
            else
            {
                renderer.PushIndent(Step);
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

    // What one level of quote indents by: the Quote style's own left indent, and what a block that
    // does not take the style (a code block, a table) is already indented by.
    const int Step = 720;

    static string Indent(int depth) =>
        (Step * (depth + 1)).ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// How many quotes enclose this one. As with a list's level, it comes from the markdown tree
    /// because a nested quote arrives as a fresh Write.
    /// </summary>
    static int ResolveDepth(QuoteBlock quoteBlock)
    {
        var depth = 0;
        for (var parent = quoteBlock.Parent; parent != null; parent = parent.Parent)
        {
            if (parent is QuoteBlock)
            {
                depth++;
            }
        }

        return depth;
    }
}
