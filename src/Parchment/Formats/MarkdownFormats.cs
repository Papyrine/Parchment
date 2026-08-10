/// <summary>
/// Honours <c>[Html]</c> / <c>[Markdown]</c> (and <c>[StringSyntax]</c>) members in the markdown
/// flow, by rewriting each of their tokens to carry the matching filter.
/// </summary>
/// <remarks>
/// <para>
/// The docx flow reads <see cref="FormatMap"/> per token at render time and turns the value into an
/// <c>HtmlToken</c> or <c>MarkdownToken</c> (<c>ScopeTreeRunner.TryResolveFormatted</c>). A markdown
/// template has no per-token dispatch to hook — the whole source goes through Fluid in one pass — so
/// the annotation is applied to the source instead, before it is parsed.
/// </para>
/// <para>
/// Rewriting at registration rather than at render keeps it off the hot path, the same way
/// <see cref="MarkdownExcelsiorTables"/> does: the rewritten source becomes the cached
/// <c>IFluidTemplate</c> and a render sees an ordinary filtered token.
/// </para>
/// <para>
/// Before bound values were escaped this was invisible — nothing was encoded, so an annotated
/// member's markdown rendered whether or not anything read the annotation. Now the annotation is
/// what separates markup from text, so it has to be read.
/// </para>
/// </remarks>
static class MarkdownFormats
{
    /// <summary>
    /// A substitution of a plain dotted path and nothing else. A token carrying a filter or an
    /// expression is left alone: its value is the filter's output rather than the member, which is
    /// the same reason the docx flow rejects that shape as PARCH010.
    /// </summary>
    static readonly Regex plainToken = new(
        @"\{\{\s*(?<path>[A-Za-z_][A-Za-z0-9_]*(?:\s*\.\s*[A-Za-z_][A-Za-z0-9_]*)*)\s*\}\}",
        RegexOptions.Compiled);

    public static string Rewrite(string markdown, FormatMap formats)
    {
        if (formats.IsEmpty)
        {
            return markdown;
        }

        return plainToken.Replace(
            markdown,
            match =>
            {
                // Fluid tolerates spaces around the dots; the map is keyed on the bare path.
                var path = match.Groups["path"].Value
                    .Replace(" ", "")
                    .Replace("\t", "");

                if (!formats.TryGet(path, out var entry))
                {
                    return match.Value;
                }

                var filter = entry.Kind switch
                {
                    FormatMapKind.Html => "html",
                    FormatMapKind.Markdown => "markdown",
                    _ => null
                };

                if (filter == null)
                {
                    return match.Value;
                }

                return $"{{{{ {path} | {filter} }}}}";
            });
    }
}
