namespace Parchment;

/// <summary>
/// Thrown from the markdown flow's value converter when a token has no markdown representation.
/// A Fluid value converter cannot see the template name, so this is caught in
/// <c>RegisteredMarkdownTemplate.Render</c> and rethrown as a <see cref="ParchmentRenderException"/>.
/// </summary>
sealed class TokenNotRenderableException(TokenValue token) :
    Exception(BuildMessage(token))
{
    static string BuildMessage(TokenValue token) =>
        $"'{token.GetType().Name}' cannot be rendered by a markdown template. It emits OpenXML " +
        "directly, which has nowhere to go in the markdown flow — liquid is rendered to markdown " +
        "text first, and only then parsed. Use MarkdownToken, HtmlToken, or plain text here, or " +
        "register the template with RegisterDocxTemplate, whose token flow substitutes OpenXML " +
        "structurally.";
}
