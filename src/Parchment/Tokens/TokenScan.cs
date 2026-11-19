/// <summary>
/// Hand-written scanner that finds Liquid token sites — <c>{{ ... }}</c> substitutions and
/// <c>{% ... %}</c> block tags — inside arbitrary text. Replaces a paragraph-splitting regex
/// that used to live on both the runtime library and the source generator, so both sides
/// agree byte-for-byte on what counts as a token. Behaviour matches the original regex
/// <c>\{\{[^{}]*?\}\}|\{%[^{%]*?%\}</c> exactly: a substitution body cannot contain
/// <c>{</c> or <c>}</c>, a block body cannot contain <c>{</c> or <c>%</c>, and the scanner
/// advances past each match without nesting.
/// </summary>
static class TokenScan
{
    public static List<TokenSite> Scan(string text)
    {
        var sites = new List<TokenSite>();
        if (text.Length < 4)
        {
            return sites;
        }

        var i = 0;
        while (i <= text.Length - 4)
        {
            var openIndex = text.IndexOf('{', i);
            if (openIndex < 0 ||
                openIndex > text.Length - 4)
            {
                break;
            }

            var second = text[openIndex + 1];
            if (second == '{' &&
                TryMatchSubstitution(text, openIndex, out var subLength))
            {
                sites.Add(new(openIndex, subLength, TokenSiteKind.Substitution));
                i = openIndex + subLength;
                continue;
            }

            if (second == '%' &&
                TryMatchBlock(text, openIndex, out var blockLength))
            {
                sites.Add(new(openIndex, blockLength, TokenSiteKind.Block));
                i = openIndex + blockLength;
                continue;
            }

            i = openIndex + 1;
        }

        return sites;
    }

    /// <summary>
    /// Returns true if any non-whitespace character in <paramref name="text"/> sits outside
    /// every site in <paramref name="sites"/>. Used to detect static text mixed in with
    /// token sites.
    /// </summary>
    public static bool HasContentOutsideSites(string text, List<TokenSite> sites)
    {
        var pos = 0;
        foreach (var site in sites)
        {
            if (HasNonWhitespace(text, pos, site.Offset))
            {
                return true;
            }
            pos = site.Offset + site.Length;
        }

        return HasNonWhitespace(text, pos, text.Length);
    }

    /// <summary>
    /// Returns true if the site at <paramref name="offset"/> is the only one in
    /// <paramref name="text"/> and everything around it is whitespace. This is what "alone in
    /// its paragraph" means for a token whose output replaces the paragraph: padding typed
    /// around the token is invisible in Word and goes with the paragraph, so it does not count
    /// as content the swap would lose. The runtime and the source generator both ask here.
    /// </summary>
    public static bool IsSolo(string text, int siteCount, int offset, int length) =>
        siteCount == 1 &&
        !HasNonWhitespace(text, 0, offset) &&
        !HasNonWhitespace(text, offset + length, text.Length);

    /// <summary>
    /// The reason an empty substitution is refused, worded once for every flow.
    /// </summary>
    public const string EmptySubstitution = "'{{ }}' names no value.";

    /// <summary>
    /// Whether the substitution <paramref name="source"/> holds nothing but whitespace control:
    /// <c>{{ }}</c> or <c>{{- -}}</c>.
    /// </summary>
    /// <remarks>
    /// Fluid 3 reads one as an empty string, as Liquid does; Fluid 2 refused it, and so does
    /// Parchment, so a token left blank by mistake fails registration rather than printing
    /// nothing. Read from the source text because the tree cannot tell it apart: Fluid parses
    /// <c>{{ empty }}</c> to the same statement.
    /// </remarks>
    public static bool IsEmptySubstitution(string source)
    {
        var start = 2;
        var end = source.Length - 2;
        if (start < end &&
            source[start] == '-')
        {
            start++;
        }

        if (start < end &&
            source[end - 1] == '-')
        {
            end--;
        }

        return !HasNonWhitespace(source, start, end);
    }

    static bool HasNonWhitespace(string text, int start, int end)
    {
        for (var k = start; k < end; k++)
        {
            if (!char.IsWhiteSpace(text[k]))
            {
                return true;
            }
        }

        return false;
    }

    static bool TryMatchSubstitution(string text, int start, out int length)
    {
        // Body must contain no `{` or `}` (mirrors regex char class `[^{}]`).
        for (var k = start + 2; k < text.Length - 1; k++)
        {
            var c = text[k];
            if (c == '}')
            {
                if (text[k + 1] == '}')
                {
                    length = k + 2 - start;
                    return true;
                }

                length = 0;
                return false;
            }

            if (c == '{')
            {
                length = 0;
                return false;
            }
        }

        length = 0;
        return false;
    }

    static bool TryMatchBlock(string text, int start, out int length)
    {
        // Body must contain no `{` or `%` (mirrors regex char class `[^{%]`).
        for (var k = start + 2; k < text.Length - 1; k++)
        {
            var c = text[k];
            if (c == '%')
            {
                if (text[k + 1] == '}')
                {
                    length = k + 2 - start;
                    return true;
                }

                length = 0;
                return false;
            }

            if (c == '{')
            {
                length = 0;
                return false;
            }
        }

        length = 0;
        return false;
    }
}
