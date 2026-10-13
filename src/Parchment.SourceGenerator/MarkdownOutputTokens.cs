/// <summary>
/// One <c>{{ ... }}</c> of a markdown template, as a read of the source text needs to see it.
/// </summary>
/// <param name="Start">Where the opening braces start in the source.</param>
/// <param name="End">The position after the closing braces.</param>
/// <param name="TrimLeft">Whether it opens with <c>{{-</c>.</param>
/// <param name="TrimRight">Whether it closes with <c>-}}</c>.</param>
/// <param name="Path">
/// The dotted member path the expression starts with, or null when it starts with something else.
/// </param>
/// <param name="IsPlain">Whether the expression is that path and nothing more.</param>
sealed record MarkdownOutputToken(int Start, int End, bool TrimLeft, bool TrimRight, string? Path, bool IsPlain);

/// <summary>
/// Finds the output tokens of a markdown template in its source text, and classifies each from the
/// expression Fluid parses it into.
/// </summary>
/// <remarks>
/// A copy of the runtime's <c>Parchment/Markdown/MarkdownOutputTokens.cs</c>, which the generator
/// cannot reference. <c>ParchmentTemplateGenerator.ValidateMarkdownExcelsiorTokens</c> has to take
/// and turn away exactly the tokens the runtime's <c>MarkdownExcelsiorTables</c> does, in every
/// spelling Fluid reads as the same token: <c>{{- Member -}}</c>, <c>{{ Owner['Member'] }}</c>, one
/// broken over several lines. Change the two together.
/// </remarks>
static class MarkdownOutputTokens
{
    public static List<MarkdownOutputToken> Scan(string markdown, FluidParser parser)
    {
        var tokens = new List<MarkdownOutputToken>();
        var position = 0;
        while (true)
        {
            var start = markdown.IndexOf("{{", position, StringComparison.Ordinal);
            if (start < 0)
            {
                break;
            }

            var close = FindClose(markdown, start + 2);
            if (close < 0)
            {
                // Unclosed, which is a parse error Fluid reports once the rewrites are done.
                break;
            }

            position = close + 2;

            var innerStart = start + 2;
            var innerEnd = close;
            var trimLeft = innerStart < innerEnd && markdown[innerStart] == '-';
            if (trimLeft)
            {
                innerStart++;
            }

            var trimRight = innerStart < innerEnd && markdown[innerEnd - 1] == '-';
            if (trimRight)
            {
                innerEnd--;
            }

            var (path, plain) = Classify(markdown.Substring(innerStart, innerEnd - innerStart), parser);
            tokens.Add(new(start, position, trimLeft, trimRight, path, plain));
        }

        return tokens;
    }

    /// <summary>
    /// Whether <paramref name="token"/> is the only content of the line or lines it sits on, and
    /// where those lines start and end. The end excludes the newline.
    /// </summary>
    public static bool IsAlone(string markdown, MarkdownOutputToken token, out int lineStart, out int lineEnd)
    {
        lineStart = 0;
        if (token.Start > 0)
        {
            lineStart = markdown.LastIndexOf('\n', token.Start - 1) + 1;
        }

        lineEnd = markdown.IndexOf('\n', token.End);
        if (lineEnd < 0)
        {
            lineEnd = markdown.Length;
        }

        return IsBlank(markdown, lineStart, token.Start) &&
               IsBlank(markdown, token.End, lineEnd);
    }

    static bool IsBlank(string markdown, int start, int end)
    {
        for (var index = start; index < end; index++)
        {
            if (!char.IsWhiteSpace(markdown[index]))
            {
                return false;
            }
        }

        return true;
    }

    // The closing braces, stepping over string literals so one holding "}}" does not end the token.
    static int FindClose(string markdown, int from)
    {
        var quote = '\0';
        for (var index = from; index < markdown.Length; index++)
        {
            var current = markdown[index];
            if (quote != '\0')
            {
                if (current == quote)
                {
                    quote = '\0';
                }

                continue;
            }

            if (current is '\'' or '"')
            {
                quote = current;
                continue;
            }

            if (current == '}' &&
                index + 1 < markdown.Length &&
                markdown[index + 1] == '}')
            {
                return index;
            }
        }

        return -1;
    }

    static (string? Path, bool IsPlain) Classify(string expression, FluidParser parser)
    {
        // Anything Fluid cannot parse fails the whole template the same way a moment later.
        if (!parser.TryParse("{{" + expression + "}}", out var template, out _) ||
            template is not Fluid.Parser.FluidTemplate { Statements: [OutputStatement output] })
        {
            return (null, false);
        }

        // A filter's value is the filter's output, and the member it started from is its input.
        var plain = true;
        var current = output.Expression;
        while (current is FilterExpression filter)
        {
            plain = false;
            current = filter.Input;
        }

        if (current is not MemberExpression member)
        {
            return (null, false);
        }

        // `Owner['Member']` is `Owner.Member` to Fluid, so the two spell one path. Any other
        // indexer ends it: what follows is an element, not a member of the model.
        var builder = new StringBuilder();
        var count = 0;
        foreach (var segment in member.Segments)
        {
            var name = SegmentNames.TryGetStaticName(segment);
            if (name == null)
            {
                break;
            }

            if (count > 0)
            {
                builder.Append('.');
            }

            builder.Append(name);
            count++;
        }

        if (count == 0)
        {
            return (null, false);
        }

        return (builder.ToString(), plain && count == member.Segments.Count);
    }
}
