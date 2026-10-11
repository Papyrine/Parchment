/// <summary>
/// Where each <c>{% for %}</c> of a markdown template binds its variable, so a read of the source
/// text can tell a token that reads a loop item from one that reads the root model.
/// </summary>
/// <remarks>
/// A copy of the runtime's <c>Parchment/Markdown/MarkdownLoopScope.cs</c>, which the generator
/// cannot reference. <c>ParchmentTemplateGenerator.ValidateMarkdownExcelsiorTokens</c> reads the
/// source the way the runtime's <c>MarkdownExcelsiorTables</c> does, and has to turn away the same
/// tokens: inside <c>{% for order in Orders %}</c> the token <c>{{ order.Lines }}</c> spells the
/// same path as a root <c>Order.Lines</c> but reads the loop item. Change the two together.
/// </remarks>
sealed class MarkdownLoopScope
{
    static readonly Regex loopTag = new(
        @"\{%-?\s*(?:for\s+(?<variable>[^\s%]+)\s+in\s|(?<end>endfor)\b)",
        RegexOptions.Compiled);

    readonly List<(int Start, int End, string Variable)> loops = [];

    public static MarkdownLoopScope Scan(string markdown)
    {
        var scope = new MarkdownLoopScope();
        var open = new Stack<(int Start, string Variable)>();
        foreach (Match match in loopTag.Matches(markdown))
        {
            if (!match.Groups["end"].Success)
            {
                open.Push((match.Index, match.Groups["variable"].Value));
                continue;
            }

            if (open.Count > 0)
            {
                var (start, variable) = open.Pop();
                scope.loops.Add((start, match.Index, variable));
            }
        }

        // An unclosed loop is a parse error, reported before this is consulted.
        foreach (var (start, variable) in open)
        {
            scope.loops.Add((start, markdown.Length, variable));
        }

        return scope;
    }

    /// <summary>
    /// Whether the token at <paramref name="position"/> whose dotted path is
    /// <paramref name="path"/> starts with a loop variable bound there. Compared by exact
    /// spelling, which is how Fluid finds a loop variable; model members are the ones matched
    /// without regard to case.
    /// </summary>
    public bool Shadows(int position, string path)
    {
        foreach (var (start, end, variable) in loops)
        {
            if (position > start &&
                position < end &&
                StartsWith(path, variable))
            {
                return true;
            }
        }

        return false;
    }

    static bool StartsWith(string path, string variable)
    {
        if (!path.StartsWith(variable, StringComparison.Ordinal))
        {
            return false;
        }

        return path.Length == variable.Length ||
               path[variable.Length] == '.';
    }
}
