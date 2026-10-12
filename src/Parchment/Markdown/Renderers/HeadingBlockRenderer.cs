class HeadingBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HeadingBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, HeadingBlock block) =>
        Write(renderer, block, new());

    /// <summary>
    /// Writes a heading into paragraph properties the caller has already given the container's
    /// share of: a list item's numbering, or a quote's indent.
    /// </summary>
    /// <remarks>
    /// A paragraph has one style, and a heading inside a list item or a quote used to be given the
    /// container's, which made it body text: absent from the navigation pane and from a table of
    /// contents. The style is what makes a heading one, so the heading keeps it and the container
    /// is stated as direct formatting instead.
    /// </remarks>
    internal static void Write(OpenXmlMarkdownRenderer renderer, HeadingBlock block, ParagraphProperties properties)
    {
        var level = Math.Clamp(block.Level + renderer.HeadingOffset, 1, 9);
        properties.ParagraphStyleId = new()
        {
            Val = ResolveStyle(block, level)
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
