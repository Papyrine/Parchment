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
/// For the same reason it is also what knows where the source is not template code at all: a
/// <c>{% raw %}</c> or <c>{% comment %}</c> region can hold anything, including what reads as a
/// token or as a loop tag, and none of it is one. See <see cref="IsLiteral"/>.
/// </para>
/// <para>
/// The source generator carries a copy (<c>Parchment.SourceGenerator/MarkdownLoopScope.cs</c>) for
/// its own read of the source, <c>ValidateMarkdownExcelsiorTokens</c>. Change the two together.
/// </para>
/// </remarks>
sealed class MarkdownLoopScope
{
    // The opening of a tag, as far as it takes to know what it is: a `for` and the variable it
    // binds, an inline comment, or the tag's name.
    static readonly Regex tag = new(
        @"\{%-?\s*(?:for\s+(?<variable>[^\s%]+)\s+in\s|(?<inline>\#)|(?<name>[a-z]+)\b)",
        RegexOptions.Compiled);

    static readonly Regex tagEnd = new(@"\G\s*-?%\}", RegexOptions.Compiled);
    static readonly Regex endRaw = new(@"\{%-?\s*endraw\s*-?%\}", RegexOptions.Compiled);
    static readonly Regex endComment = new(@"\{%-?\s*endcomment\s*-?%\}", RegexOptions.Compiled);

    readonly List<(int Start, int End, string Variable)> loops = [];
    readonly List<(int Start, int End)> literals = [];

    public static MarkdownLoopScope Scan(string markdown)
    {
        var scope = new MarkdownLoopScope();

        // Every open block that can hold an `{% else %}`, so that an else is given to the block it
        // belongs to. A null variable is a block that binds nothing from here on: an if, unless or
        // case, or a loop already past its else.
        var open = new Stack<(int Start, string? Variable)>();
        var position = 0;
        while (true)
        {
            var match = tag.Match(markdown, position);
            if (!match.Success)
            {
                break;
            }

            position = match.Index + match.Length;
            if (match.Groups["variable"].Success)
            {
                open.Push((match.Index, match.Groups["variable"].Value));
                continue;
            }

            if (match.Groups["inline"].Success)
            {
                // `{% # ... %}` is a comment that ends with its own tag.
                var close = markdown.IndexOf("%}", position, StringComparison.Ordinal);
                position = scope.AddLiteral(match.Index, close, 2, markdown);
                continue;
            }

            switch (match.Groups["name"].Value)
            {
                case "raw":
                {
                    // Fluid reads the region from the end of the opening tag, so a tag that does
                    // not end straight after the name opens nothing.
                    var opened = tagEnd.Match(markdown, position);
                    if (opened.Success)
                    {
                        var close = endRaw.Match(markdown, opened.Index + opened.Length);
                        position = scope.AddLiteral(match.Index, close.Index, close.Length, markdown);
                    }

                    break;
                }

                case "comment":
                {
                    var close = endComment.Match(markdown, position);
                    position = scope.AddLiteral(match.Index, close.Index, close.Length, markdown);
                    break;
                }

                case "if":
                case "unless":
                case "case":
                    open.Push((match.Index, null));
                    break;

                case "else":
                    // The else of a loop renders when there was nothing to iterate, so the variable
                    // is not bound in it: the loop's extent stops here rather than at its endfor.
                    if (open.Count > 0 &&
                        open.Peek().Variable != null)
                    {
                        var (start, variable) = open.Pop();
                        scope.loops.Add((start, match.Index, variable!));
                        open.Push((start, null));
                    }

                    break;

                case "endfor":
                case "endif":
                case "endunless":
                case "endcase":
                    if (open.Count > 0)
                    {
                        var (start, variable) = open.Pop();
                        if (variable != null)
                        {
                            scope.loops.Add((start, match.Index, variable));
                        }
                    }

                    break;
            }
        }

        // An unclosed loop is a parse error Fluid reports once the rewrites are done.
        foreach (var (start, variable) in open)
        {
            if (variable != null)
            {
                scope.loops.Add((start, markdown.Length, variable));
            }
        }

        return scope;
    }

    // Records a region that ends with the closing tag found at close, and returns where the scan
    // resumes. An unclosed region runs to the end of the source, which is a parse error too. A
    // failed Match reports an index of 0, which no closing tag can have, since it follows an
    // opening one.
    int AddLiteral(int start, int close, int closeLength, string markdown)
    {
        var end = markdown.Length;
        if (close > 0)
        {
            end = close + closeLength;
        }

        literals.Add((start, end));
        return end;
    }

    /// <summary>
    /// Whether <paramref name="position"/> is inside <c>{% raw %}</c>, <c>{% comment %}</c> or an
    /// inline <c>{% # %}</c> comment. What looks like a token there is text Fluid prints as written
    /// or drops, never a read of the model, so it is neither rewritten nor judged.
    /// </summary>
    public bool IsLiteral(int position)
    {
        foreach (var (start, end) in literals)
        {
            if (position >= start &&
                position < end)
            {
                return true;
            }
        }

        return false;
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
