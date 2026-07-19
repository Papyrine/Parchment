class LiteralInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, LiteralInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, LiteralInline inline)
    {
        var content = inline.Content.AsSpan();
        if (content.Length == 0)
        {
            return;
        }

        var text = XmlCharSanitizer.Strip(content).ToString();
        renderer.AddRun(BuildRun(text));
    }

    /// <summary>
    /// Builds the run for a literal, turning any tab into a real Word tab.
    /// </summary>
    /// <remarks>
    /// A tab left inside <c>&lt;w:t&gt;</c> is not a Word tab — Word wants a <c>&lt;w:tab/&gt;</c>
    /// sibling of the text, and renders the raw character as ordinary whitespace instead. There was
    /// otherwise no way to ask for one: html folds tabs to spaces, so nothing in either flow reached
    /// <c>&lt;w:tab/&gt;</c> at all, and tabs are load-bearing layout in a formal document.
    /// </remarks>
    internal static Run BuildRun(string text)
    {
        var tab = text.IndexOf('\t');
        if (tab < 0)
        {
            return new(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
        }

        var run = new Run();
        var start = 0;
        while (tab >= 0)
        {
            if (tab > start)
            {
                run.Append(
                    new Text(text[start..tab])
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
            }

            run.Append(new TabChar());
            start = tab + 1;
            tab = text.IndexOf('\t', start);
        }

        if (start < text.Length)
        {
            run.Append(
                new Text(text[start..])
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
        }

        return run;
    }
}
