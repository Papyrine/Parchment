namespace Parchment;

/// <summary>
/// Renders a <see cref="TokenValue"/> as markdown source for the markdown flow.
/// </summary>
/// <remarks>
/// The docx flow substitutes tokens structurally — <c>ScopeTreeRunner</c> replaces the host
/// paragraph with whatever the token produces. The markdown flow has no OpenXML to replace at
/// substitution time: liquid renders to markdown text first, and only then is that text parsed.
/// So a value has to become text here, and the ones that only exist to emit OpenXML have no text
/// form at all.
/// </remarks>
static class TokenMarkdown
{
    public static string Render(TokenValue token) =>
        token switch
        {
            TextToken text => text.Value,
            // Markdown source written into a markdown template is already what it will be parsed
            // as, so it passes straight through.
            MarkdownToken markdown => markdown.Source,
            // Html is not: written into the source it would be classified by Markdig rather than
            // converted, which is a different answer. So it takes the same marker route the html
            // filter does and is converted once the parse is done — see MarkdownHtmlBlocks.
            HtmlToken html => MarkdownHtmlBlocks.Register(html.Source),
            _ => throw new TokenNotRenderableException(token)
        };
}
