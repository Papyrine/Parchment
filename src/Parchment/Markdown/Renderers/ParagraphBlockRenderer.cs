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

        if (TableOfContents.TryWrite(renderer, block, properties))
        {
            return;
        }

        if (SectionBreaks.TryWrite(renderer, block, properties))
        {
            return;
        }

        var bookmark = MarkdownBookmark.Resolve(block);
        var bookmarkId = bookmark == null ? null : renderer.AddBookmarkStart(bookmark);
        renderer.WriteLeafInline(block);
        if (bookmarkId != null)
        {
            renderer.AddBookmarkEnd(bookmarkId);
        }

        renderer.FlushParagraph(properties);
    }
}
