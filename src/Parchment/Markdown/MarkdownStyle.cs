using Markdig.Renderers.Html;

/// <summary>
/// Reads the style name from a <c>{.StyleName}</c> generic attribute.
/// </summary>
/// <remarks>
/// The target is decided by whatever is being rendered rather than by the name — a name on a table
/// becomes a table style, on a run a character style, on a paragraph a paragraph style. The name is
/// not checked against the style source: Word resolves latent built-in styles that never appear in
/// styles.xml, so an unknown name is not evidence of a mistake.
/// <para>
/// Markdig's generic-attribute syntax also carries an id (<c>{#id}</c>), key=value properties and
/// further classes. A Word style is a single name, so only the first class maps onto anything and
/// the rest are dropped. That keeps markdown written for both html and Word portable rather than
/// rejecting it, and it is the documented behaviour — see the generic attributes section of the
/// readme.
/// </para>
/// </remarks>
static class MarkdownStyle
{
    public static string? Resolve(IMarkdownObject node) =>
        node.TryGetAttributes()?.Classes?.FirstOrDefault();
}
