namespace Parchment;

/// <summary>
/// Renders a <see cref="TokenValue"/> as markdown source for the markdown flow.
/// </summary>
/// <remarks>
/// The docx flow substitutes tokens structurally — <c>ScopeTreeRunner</c> replaces the host
/// paragraph with whatever the token produces. The markdown flow has no OpenXML to replace at
/// substitution time: liquid renders to markdown text first, and only then is that text parsed.
/// So a token has to become text here, and the ones that only exist to emit OpenXML have no text
/// form at all.
/// </remarks>
static class TokenMarkdown
{
    public static string Render(TokenValue token) =>
        token switch
        {
            TextToken text => text.Value,
            // Both are already source the markdown parser handles: markdown directly, and html via
            // an html block.
            MarkdownToken markdown => markdown.Source,
            HtmlToken html => html.Source,
            _ => throw new TokenNotRenderableException(token)
        };
}
