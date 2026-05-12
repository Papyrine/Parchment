/// <summary>
/// SG-side mirror of <c>Parchment.Tokens.BlockTagParser</c>. The SG project can't reference the
/// runtime assembly (different target frameworks; SG ships as a netstandard2.0 analyzer), so the
/// parser is duplicated here. Same semantics as the original
/// <c>^\{%\s*(?&lt;tag&gt;\w+)(?:\s+(?&lt;expr&gt;.*?))?\s*%\}$</c> regex, span-based, allocation-free.
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

        var inner = span.Slice(2, span.Length - 4).Trim();
        if (inner.IsEmpty)
        {
            return false;
        }

        var tagLength = 0;
        while (tagLength < inner.Length &&
               IsWord(inner[tagLength]))
        {
            tagLength++;
        }

        if (tagLength == 0)
        {
            return false;
        }

        tag = inner[..tagLength];
        var rest = inner[tagLength..];

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
