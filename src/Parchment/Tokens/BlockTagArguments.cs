/// <summary>
/// The tags that take no arguments: <c>else</c> and every closing tag. Liquid rejects anything
/// after the name, so <c>{% else if X %}</c> is an error rather than a branch with a condition.
/// Compiled into both the runtime library and the source generator so the four validators word
/// the rejection the same way.
/// </summary>
static class BlockTagArguments
{
    const string useElsif = "Use '{% elsif condition %}'.";

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
    /// The first such problem among the block tags of <paramref name="text"/>. The markdown flow
    /// asks only once Fluid has refused the template, to say which tag it refused and why — Fluid
    /// names the tag it was expecting instead, which reads as a missing <c>{% endif %}</c>.
    /// </summary>
    public static string? FirstProblem(string text)
    {
        foreach (var site in TokenScan.Scan(text))
        {
            if (site.Kind == TokenSiteKind.Substitution)
            {
                continue;
            }

            var source = text.Substring(site.Offset, site.Length);
            if (!BlockTagParser.TryParse(source, out var tag, out var expression))
            {
                continue;
            }

            string? argument = null;
            if (!expression.IsEmpty)
            {
                argument = expression.ToString();
            }

            var problem = Problem(tag.ToString(), argument);
            if (problem != null)
            {
                return problem;
            }
        }

        return null;
    }

    /// <summary>
    /// The message for a markdown template Fluid could not parse, led by the tag problem when
    /// there is one.
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

    static bool IsIf(string expression) =>
        expression.StartsWith("if", StringComparison.Ordinal) &&
        (expression.Length == 2 || char.IsWhiteSpace(expression[2]));
}
