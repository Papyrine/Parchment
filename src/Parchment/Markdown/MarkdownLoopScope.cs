/// <summary>
/// Where each <c>{% for %}</c> of a markdown template binds its variable, so the registration-time
/// rewrites can tell a token that reads a loop item from one that reads the root model.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="MarkdownExcelsiorTables"/> and <see cref="MarkdownFormats"/> rewrite tokens in the
/// source text, before Fluid has parsed it, and match them to maps keyed on dotted paths from the
/// root model. Inside <c>{% for customer in Customers %}</c> the token <c>{{ customer.Notes }}</c>
/// spells the same path as a root <c>Customer.Notes</c> but reads the loop item, so it has to be
/// left for Fluid. This is the markdown counterpart of <see cref="LoopScope"/>, working from
/// positions in the text because there is no tree yet.
/// </para>
/// <para>
/// The source generator carries a copy (<c>Parchment.SourceGenerator/MarkdownLoopScope.cs</c>) for
/// its own read of the source, <c>ValidateMarkdownExcelsiorTokens</c>. Change the two together.
/// </para>
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

        // An unclosed loop is a parse error Fluid reports once the rewrites are done.
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
