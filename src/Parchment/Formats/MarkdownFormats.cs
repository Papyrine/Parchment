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
    /// Appends the filter to each token that is a plain path to an annotated member, in whichever
    /// spelling Fluid reads as one (see <see cref="MarkdownOutputTokens"/>). A token carrying a
    /// filter or an expression is left alone: its value is the filter's output rather than the
    /// member, which is the same reason the docx flow rejects that shape as PARCH010.
    /// </summary>
    public static string Rewrite(string markdown, FormatMap formats)
    {
        if (formats.IsEmpty)
        {
            return markdown;
        }

        var loops = MarkdownLoopScope.Scan(markdown);
        StringBuilder? builder = null;
        var copied = 0;
        foreach (var token in MarkdownOutputTokens.Scan(markdown, SharedFluid.Parser))
        {
            // A token that starts with a loop variable reads the loop item, whatever root
            // path it happens to spell, so the root member's annotation is not its own.
            // Nor is what sits in a raw or comment region a token at all.
            if (token.Path == null ||
                !token.IsPlain ||
                loops.IsLiteral(token.Start) ||
                loops.Shadows(token.Start, token.Path) ||
                !formats.TryGet(token.Path, out var entry))
            {
                continue;
            }

            var filter = entry.Kind switch
            {
                FormatMapKind.Html => "html",
                FormatMapKind.Markdown => "markdown",
                _ => null
            };

            if (filter == null)
            {
                continue;
            }

            // The dashes stay where they were, so the token still trims what it trimmed.
            builder ??= new();
            builder.Append(markdown, copied, token.Start - copied);
            builder.Append(token.TrimLeft ? "{{- " : "{{ ");
            builder.Append(token.Path);
            builder.Append(" | ");
            builder.Append(filter);
            builder.Append(token.TrimRight ? " -}}" : " }}");
            copied = token.End;
        }

        if (builder == null)
        {
            return markdown;
        }

        builder.Append(markdown, copied, markdown.Length - copied);
        return builder.ToString();
    }
}
