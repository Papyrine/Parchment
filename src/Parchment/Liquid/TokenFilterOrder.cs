/// <summary>
/// Finds a filter applied to the result of a filter that produces a token.
/// </summary>
/// <remarks>
/// <para>
/// A filter that produces Word content hands back a token rather than text, so nothing after it in
/// the chain has text to work on. In a docx template <c>markdown</c>, <c>html</c>,
/// <c>bullet_list</c> and <c>numbered_list</c> each return a <c>TokenValue</c>, and the next filter
/// stringifies it: <c>{{ Text | markdown | upcase }}</c> printed <c>PARCHMENT.MARKDOWNTOKEN</c>. In
/// a markdown template only <c>html</c> does — it parks the value and writes a placeholder, which
/// <c>upcase</c> turned into a marker nothing recognises. The other three write markdown source
/// there, and source is text a later filter can still transform.
/// </para>
/// <para>
/// A filter chain is only ever the whole of an output or an assign value. Fluid refuses one in a
/// condition, a loop source or a <c>case</c>, and grouping with parentheses is an option neither
/// parser turns on, so one cannot sit inside an argument either. Walking the inputs from the
/// outermost filter therefore sees every link.
/// </para>
/// <para>
/// This file is compiled into the source generator as well (linked in its csproj), so the four
/// validators read one definition of which filters produce a token. A filter added to
/// <c>Filters</c> that returns a <c>TokenValue</c> belongs in <see cref="ProducesToken"/>.
/// </para>
/// </remarks>
static class TokenFilterOrder
{
    // Ordinal, as Fluid resolves a filter by its exact name: `| Markdown` is an unknown filter.
    static bool ProducesToken(string filter, bool markdownFlow)
    {
        if (markdownFlow)
        {
            return filter == "html";
        }

        return filter is "markdown" or "html" or "bullet_list" or "numbered_list";
    }

    /// <summary>
    /// The leftmost filter of <paramref name="expression"/> that follows a token-producing one.
    /// </summary>
    public static bool TryFind(
        Expression expression,
        bool markdownFlow,
        [NotNullWhen(true)] out string? tokenFilter,
        [NotNullWhen(true)] out string? following)
    {
        tokenFilter = null;
        following = null;
        var current = expression;
        while (current is FilterExpression
               {
                   Input: FilterExpression input
               } filter)
        {
            // Not returned on the first hit: the walk runs right to left, and the one to report is
            // the first a reader meets.
            if (ProducesToken(input.Name, markdownFlow))
            {
                tokenFilter = input.Name;
                following = filter.Name;
            }

            current = input;
        }

        return tokenFilter != null;
    }

    /// <summary>
    /// What is wrong and what to do about it, in the words both the registration exception and
    /// PARCH037 use.
    /// </summary>
    public static string Message(string tokenFilter, string following) =>
        $"Filter '{following}' is applied after '{tokenFilter}', which has to be the last filter of its token. '{tokenFilter}' produces Word content rather than text, so a filter after it has no text to work on. Move '{following}' before '{tokenFilter}'.";

    /// <summary>
    /// The chain as written, without filter arguments: <c>Body | html | upcase</c>. For the
    /// markdown flow, where a validator holds the parsed template and not the token's source.
    /// </summary>
    public static string Describe(Expression expression)
    {
        var filters = new List<string>();
        var current = expression;
        while (current is FilterExpression filter)
        {
            filters.Add(filter.Name);
            current = filter.Input;
        }

        var builder = new StringBuilder(DescribeInput(current));
        for (var index = filters.Count - 1; index >= 0; index--)
        {
            builder.Append(" | ");
            builder.Append(filters[index]);
        }

        return builder.ToString();
    }

    static string DescribeInput(Expression expression)
    {
        if (expression is not MemberExpression member)
        {
            return "...";
        }

        var names = new List<string>(member.Segments.Count);
        foreach (var segment in member.Segments)
        {
            if (segment is not IdentifierSegment identifier)
            {
                return "...";
            }

            names.Add(identifier.Identifier);
        }

        return string.Join(".", names);
    }
}
