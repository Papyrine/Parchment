static class RendererInlineExtensions
{
    public static void WriteLeafInline(this OpenXmlMarkdownRenderer renderer, LeafBlock block)
    {
        if (block.Inline != null)
        {
            renderer.WriteChildren(block.Inline);
        }

        // Inline html formatting ends with the block it was opened in, closed or not - the scope a
        // browser gives an unclosed inline tag inside a <p>. Every block with inline content comes
        // through here (a table cell's is a paragraph), so this is the one place that bounds it.
        renderer.ClearInlineHtmlFormats();
    }

    public static void WriteChildren(this OpenXmlMarkdownRenderer renderer, ContainerInline inline)
    {
        var child = inline.FirstChild;
        while (child != null)
        {
            renderer.Render(child);
            child = child.NextSibling;
        }
    }
}