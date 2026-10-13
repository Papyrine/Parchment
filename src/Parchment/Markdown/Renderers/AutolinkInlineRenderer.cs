class AutolinkInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, AutolinkInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, AutolinkInline inline)
    {
        var text = Unescape(inline.Url, out var closedByValue);
        var url = text;
        if (inline.IsEmail)
        {
            url = $"mailto:{text}";
        }

        // A url no relationship can point at stays the text it was written as.
        if (RelationshipIds.TryHyperlink(renderer.Part, url, out var relId))
        {
            var run = new Run(
                new RunProperties(
                    new RunStyle
                    {
                        Val = MarkdownStyle.Resolve(inline) ?? "Hyperlink"
                    }),
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            // The run goes in already wrapped, so the inline html open around it ("<b><url></b>") is
            // applied here: AddRun formats a Run, and what it is handed is the Hyperlink. A classed span
            // replaces the link's style, as it does for the runs of a "[text](url)".
            renderer.ApplyInlineHtmlFormats(run);
            renderer.AddRun(
                new Hyperlink(run)
                {
                    Id = relId
                });
        }
        else
        {
            renderer.AddRun(LiteralInlineRenderer.BuildRun(text));
        }

        // The > that closed the autolink belonged to the value, so it is text the value is owed
        // rather than syntax to swallow.
        if (closedByValue)
        {
            renderer.AddRun(LiteralInlineRenderer.BuildRun(">"));
        }
    }

    // An autolink is read verbatim, so the escapes MarkdownEncoder wrote into a bound part of the
    // url are still in it — see BoundAutoLinkParser, which does this for the bare form. A file: url
    // is left as written: a backslash is a path separator there, and \\server\share is not an
    // escaped backslash.
    static string Unescape(string url, out bool closedByValue)
    {
        closedByValue = false;
        if (url.AsSpan().IndexOfAny('\\', '&') < 0 ||
            url.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
        {
            return url;
        }

        // A value holding a > closes the autolink with it, leaving that character's escape as the
        // last thing inside.
        var span = url.AsSpan();
        if (MarkdownEncoder.EndsWithDanglingEscape(span))
        {
            closedByValue = true;
            span = span[..^1];
        }

        return MarkdownEncoder.Unescape(span);
    }
}
