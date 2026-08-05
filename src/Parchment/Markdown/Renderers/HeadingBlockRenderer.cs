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

        // The bookmark wraps the heading text inside the paragraph rather than sitting beside it,
        // which is where Word puts its own _Toc bookmarks — a PAGEREF to a bookmark that spans no
        // content has nothing to resolve a page from.
        var bookmark = MarkdownBookmark.Resolve(block);
        var bookmarkId = bookmark == null ? null : renderer.AddBookmarkStart(bookmark);
        renderer.WriteLeafInline(block);
        if (bookmarkId != null)
        {
            renderer.AddBookmarkEnd(bookmarkId);
        }

        renderer.FlushParagraph(properties);
    }

    static string ResolveStyle(HeadingBlock block, int level) =>
        MarkdownStyle.Resolve(block) ?? $"Heading{level}";
}
