namespace Parchment;

/// <summary>
/// Renders a <see cref="TokenValue"/> as markdown source for the markdown flow.
/// </summary>
/// <remarks>
/// The docx flow substitutes tokens structurally — <c>ScopeTreeRunner</c> replaces the host
/// paragraph with whatever the token produces. The markdown flow has no OpenXML to replace at
/// substitution time: liquid renders to markdown text first, and only then is that text parsed.
/// So a value that has no text form is parked against a marker instead, and swapped for what it
/// produces once the parse is done — see <see cref="MarkdownTokenBlocks"/>. What is left with no
/// answer at all is <see cref="MutateToken"/>, which mutates the paragraph a docx template already
/// had; markdown has no such paragraph to hand it.
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
            // filter does and is converted once the parse is done — see MarkdownTokenBlocks.
            HtmlToken html => MarkdownTokenBlocks.Register(html),
            // OpenXML has no markdown form at all, so it takes the same route for a stronger
            // version of the same reason: there is nothing to write, but by the time the marker is
            // swapped there is a document to emit into.
            OpenXmlToken openXml => MarkdownTokenBlocks.Register(openXml),
            _ => throw new TokenNotRenderableException(token)
        };
}
