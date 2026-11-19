/// <summary>
/// The tags that take no arguments: <c>else</c> and every closing tag. Liquid rejects anything
/// after the name, so <c>{% else if X %}</c> is an error rather than a branch with a condition.
/// Compiled into both the runtime library and the source generator so the four validators word
/// the rejection the same way. The markdown flow also asks here about an empty <c>{{ }}</c>, the
/// other mistake Fluid 3 parses and Parchment refuses, and every flow asks about a <c>for</c>
/// whose <c>reversed</c> is not where liquid reads it, to say how to write the tag.
/// </summary>
static class BlockTagArguments
{
    const string useElsif = "Use '{% elsif condition %}'.";

    const string reversed = "reversed";

    static readonly string[] modifiers = ["limit", "offset"];

    /// <summary>
    /// The reason <paramref name="tag"/> cannot carry <paramref name="expression"/>, or null when
    /// it can.
    /// </summary>
    public static string? Problem(string tag, string? expression)
    {
        if (tag == "elseif")
        {
            return $"'{{% elseif %}}' is not a liquid tag. {useElsif}";
        }

        if (expression == null)
        {
            return null;
        }

        if (tag == "for")
        {
            var moved = MoveReversed(expression);
            if (moved == null)
            {
                return null;
            }

            return $"'reversed' has to come straight after the loop source, ahead of limit and offset: '{{% for {moved} %}}'.";
        }

        if (tag != "else" &&
            !tag.StartsWith("end", StringComparison.Ordinal))
        {
            return null;
        }

        if (tag == "else" &&
            IsIf(expression))
        {
            return $"'{{% else if %}}' is not a liquid tag. {useElsif}";
        }

        return $"'{{% {tag} %}}' takes no arguments.";
    }

    /// <summary>
    /// The first such problem among the block tags of <paramref name="text"/>, or an empty
    /// <c>{{ }}</c> (see <see cref="TokenScan.IsEmptySubstitution"/>), outside the
    /// <c>{% raw %}</c> and comment regions Fluid does not read as template code.
    /// </summary>
    public static string? FirstProblem(string text)
    {
        // Only built once a tag has a problem, which most templates never have.
        MarkdownLoopScope? scope = null;
        foreach (var site in TokenScan.Scan(text))
        {
            var problem = SiteProblem(text.Substring(site.Offset, site.Length), site.Kind);
            if (problem == null)
            {
                continue;
            }

            scope ??= MarkdownLoopScope.Scan(text);
            if (!scope.IsLiteral(site.Offset))
            {
                return problem;
            }
        }

        return null;
    }

    /// <summary>
    /// The message for a markdown template Fluid could not parse, led by the tag problem when
    /// there is one: Fluid names the tag it was expecting instead, which reads as a missing
    /// <c>{% endif %}</c>.
    /// </summary>
    public static string MarkdownParseFailure(string markdown, string error)
    {
        var problem = FirstProblem(markdown);
        if (problem == null)
        {
            return $"Failed to parse markdown as a liquid template: {error}";
        }

        return $"Failed to parse markdown as a liquid template: {problem} {error}";
    }

    /// <summary>
    /// The message for a markdown template Fluid parsed although one of its tags has a problem,
    /// or null when none has.
    /// </summary>
    /// <remarks>
    /// Fluid 3 reads an argument on <c>else</c> or a closing tag and drops it, so
    /// <c>{% else if X %}</c> parses as a plain <c>else</c> and its branch renders whenever the
    /// <c>if</c> fails. Liquid rejects the tag, and so does the docx flow, which scans its own.
    /// An empty <c>{{ }}</c> is refused with them, as the docx flow refuses it.
    /// </remarks>
    public static string? MarkdownProblem(string markdown)
    {
        var problem = FirstProblem(markdown);
        if (problem == null)
        {
            return null;
        }

        return $"Failed to parse markdown as a liquid template: {problem}";
    }

    /// <summary>
    /// The loop expression of a <c>for</c> with a <c>reversed</c> written after <c>limit</c> or
    /// <c>offset</c> moved to straight after the loop source, or null when it is not written there.
    /// </summary>
    /// <remarks>
    /// Liquid reads <c>reversed</c> only in that place, and Fluid 3 follows it, refusing the tag
    /// with "Valid attributes are limit and offset", which does not say what to change. Fluid 2
    /// took it anywhere, so a template written against it may have it last.
    /// </remarks>
    static string? MoveReversed(string expression)
    {
        var firstModifier = -1;
        var position = 0;
        while (NextWord(expression, ref position, out var start))
        {
            if (firstModifier < 0)
            {
                if (IsModifier(expression, start))
                {
                    firstModifier = start;
                }

                continue;
            }

            // A `reversed` straight after a colon is a modifier's value, a variable of that name.
            if (position - start != reversed.Length ||
                string.CompareOrdinal(expression, start, reversed, 0, reversed.Length) != 0 ||
                PrecededByColon(expression, start))
            {
                continue;
            }

            var removeFrom = start;
            while (removeFrom > 0 &&
                   char.IsWhiteSpace(expression[removeFrom - 1]))
            {
                removeFrom--;
            }

            // The modifier comes first, so taking the word out leaves it where it was.
            return expression
                .Remove(removeFrom, position - removeFrom)
                .Insert(firstModifier, $"{reversed} ");
        }

        return null;
    }

    static bool NextWord(string text, ref int position, out int start)
    {
        while (position < text.Length &&
               char.IsWhiteSpace(text[position]))
        {
            position++;
        }

        start = position;
        while (position < text.Length &&
               !char.IsWhiteSpace(text[position]))
        {
            position++;
        }

        return position > start;
    }

    // Whether the word at start names a modifier: `limit` or `offset` followed, past any
    // whitespace, by its colon.
    static bool IsModifier(string text, int start)
    {
        foreach (var name in modifiers)
        {
            if (string.CompareOrdinal(text, start, name, 0, name.Length) != 0)
            {
                continue;
            }

            var index = start + name.Length;
            while (index < text.Length &&
                   char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            if (index < text.Length &&
                text[index] == ':')
            {
                return true;
            }
        }

        return false;
    }

    static bool PrecededByColon(string text, int start)
    {
        var index = start - 1;
        while (index >= 0 &&
               char.IsWhiteSpace(text[index]))
        {
            index--;
        }

        return index >= 0 &&
               text[index] == ':';
    }

    static string? SiteProblem(string source, TokenSiteKind kind)
    {
        if (kind == TokenSiteKind.Substitution)
        {
            if (TokenScan.IsEmptySubstitution(source))
            {
                return TokenScan.EmptySubstitution;
            }

            return null;
        }

        if (!BlockTagParser.TryParse(source, out var tag, out var expression))
        {
            return null;
        }

        string? argument = null;
        if (!expression.IsEmpty)
        {
            argument = expression.ToString();
        }

        return Problem(tag.ToString(), argument);
    }

    static bool IsIf(string expression) =>
        expression.StartsWith("if", StringComparison.Ordinal) &&
        (expression.Length == 2 || char.IsWhiteSpace(expression[2]));
}
