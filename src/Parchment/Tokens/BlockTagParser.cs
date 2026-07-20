/// <summary>
/// Parses a block-tag source string of the form <c>{% tag expr %}</c> into its two components.
/// Replaces the equivalent regex (<c>^\{%\s*(?&lt;tag&gt;\w+)(?:\s+(?&lt;expr&gt;.*?))?\s*%\}$</c>)
/// with a span-based hand-roll — same semantics, zero allocations, no Regex engine startup.
/// </summary>
static class BlockTagParser
{
    public static bool TryParse(
        string source,
        out CharSpan tag,
        out CharSpan expression)
    {
        tag = default;
        expression = default;

        var span = source.AsSpan();
        if (span.Length < 4 ||
            span[0] != '{' || span[1] != '%' ||
            span[^2] != '%' || span[^1] != '}')
        {
            return false;
        }

        var inner = span[2..^2];

        // Whitespace control. The hyphens in `{%-` and `-%}` belong to the delimiter, not the tag,
        // so they come off before the tag is read. They have to be adjacent to the delimiter to
        // count, which is why this runs before the trim — `{% - for %}` stays a parse error.
        // Markdown templates go through Fluid's own parser, which handles this; the docx flow scans
        // tags itself and used to reject `{%- for row in Rows %}` as a malformed block tag.
        if (inner.Length > 0 &&
            inner[0] == '-')
        {
            inner = inner[1..];
        }

        if (inner.Length > 0 &&
            inner[^1] == '-')
        {
            inner = inner[..^1];
        }

        inner = inner.Trim();
        if (inner.IsEmpty)
        {
            return false;
        }

        // Tag is one or more `\w` characters — letters, digits, underscore.
        var tagLength = 0;
        while (tagLength < inner.Length && IsWord(inner[tagLength]))
        {
            tagLength++;
        }

        if (tagLength == 0)
        {
            return false;
        }

        tag = inner[..tagLength];
        var rest = inner[tagLength..];

        // After the tag, either end-of-content or whitespace + expression. The original regex
        // required `\s+` between tag and expr, so non-whitespace immediately after tag word chars
        // (e.g. `{% for[x] %}`) is a parse failure — same here.
        if (rest.IsEmpty)
        {
            return true;
        }

        if (!char.IsWhiteSpace(rest[0]))
        {
            tag = default;
            return false;
        }

        expression = rest.TrimStart();
        return true;
    }

    static bool IsWord(char c) =>
        char.IsLetterOrDigit(c) || c == '_';
}
