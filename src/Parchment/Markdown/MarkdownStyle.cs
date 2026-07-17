using Markdig.Renderers.Html;

/// <summary>
/// Reads the style name from a <c>{.StyleName}</c> generic attribute.
/// </summary>
/// <remarks>
/// The target is decided by whatever is being rendered rather than by the name — a name on a table
/// becomes a table style, on a run a character style, on a paragraph a paragraph style. The name is
/// not checked against the style source: Word resolves latent built-in styles that never appear in
/// styles.xml, so an unknown name is not evidence of a mistake.
/// </remarks>
static class MarkdownStyle
{
    public static string? Resolve(IMarkdownObject node) =>
        node.TryGetAttributes()?.Classes?.FirstOrDefault();
}
