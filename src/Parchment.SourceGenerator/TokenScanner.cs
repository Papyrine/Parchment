public static class TokenScanner
{
    static readonly FluidParser parser = new();

    // containers runs parallel to paragraphs: the id of the element each paragraph sits in, copied
    // onto its tokens. Omitted, every token reports container 0.
    public static List<Token> Scan(IEnumerable<string> paragraphs, IReadOnlyList<int>? containers = null)
    {
        var result = new List<Token>();
        var index = -1;
        foreach (var paragraph in paragraphs)
        {
            index++;
            var sites = TokenScan.Scan(paragraph);
            if (sites.Count == 0)
            {
                continue;
            }

            var hasOtherContent = TokenScan.HasContentOutsideSites(paragraph, sites);
            var container = 0;
            if (containers != null)
            {
                container = containers[index];
            }

            foreach (var site in sites)
            {
                var source = paragraph.Substring(site.Offset, site.Length);
                if (site.Kind == TokenSiteKind.Substitution)
                {
                    // A sibling token is other content to a substitution: the runtime asks the
                    // same TokenScan.IsSolo before it swaps a paragraph for a token's output.
                    var solo = TokenScan.IsSolo(paragraph, sites.Count, site.Offset, site.Length);
                    result.Add(ParseSubstitution(source, !solo) with { Container = container });
                }
                else
                {
                    result.Add(ParseBlockTag(source, hasOtherContent) with { Container = container });
                }
            }
        }

        return result;
    }

    static Token ParseSubstitution(string source, bool hasOtherContent)
    {
        if (!parser.TryParse(source, out var template, out _))
        {
            return new(TokenKind.Substitution, source, [], null, hasOtherContent);
        }

        var references = IdentifierVisitor.Collect(template);
        var isPlain = IsPlainMemberAccess(template);
        var token = new Token(TokenKind.Substitution, source, references, null, hasOtherContent, isPlain);

        // The docx form of the rule, where every one of these filters produces a token: this
        // scanner only ever reads a docx template. Lockstep with the runtime's
        // TokenScanner.RequireTokenFilterLast.
        var statements = ((FluidTemplate) template).Statements;
        if (statements.Count > 0 &&
            statements[0] is OutputStatement output &&
            TokenFilterOrder.TryFind(output.Expression, markdownFlow: false, out var tokenFilter, out var following))
        {
            return token with
            {
                TokenFilter = tokenFilter,
                FilterAfterTokenFilter = following
            };
        }

        return token;
    }

    // A "plain" substitution is `{{ a.b.c }}` — just an OutputStatement wrapping a MemberExpression,
    // no filter chain, no arithmetic, no literal. The Excelsior dispatch path requires this shape
    // because it walks the model object directly instead of going through Fluid evaluation.
    static bool IsPlainMemberAccess(IFluidTemplate template)
    {
        var statements = ((FluidTemplate) template).Statements;
        return statements.Count == 1 &&
               statements[0] is
                   OutputStatement
                   {
                       Expression: MemberExpression
                   };
    }

    static Token ParseBlockTag(string source, bool hasOtherContent)
    {
        if (!BlockTagParser.TryParse(source, out var tag, out var expressionSpan))
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        var expression = expressionSpan.IsEmpty ? null : expressionSpan.ToString();

        switch (tag)
        {
            case "for":
                return BuildForTag(source, expression, hasOtherContent);

            case "endfor":
                return BuildArgumentless(TokenKind.ForClose, "endfor", source, expression, hasOtherContent);

            case "if":
                return BuildConditional(TokenKind.IfOpen, source, expression, hasOtherContent);

            case "elsif":
                return BuildConditional(TokenKind.ElsIf, source, expression, hasOtherContent);

            case "else":
                return BuildArgumentless(TokenKind.Else, "else", source, expression, hasOtherContent);

            case "endif":
                return BuildArgumentless(TokenKind.IfClose, "endif", source, expression, hasOtherContent);

            case "elseif":
                return BuildArgumentless(TokenKind.UnknownBlock, "elseif", source, expression, hasOtherContent);

            default:
                return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }
    }

    // Lockstep with the runtime TokenScanner.RequireNoArguments: `{% else if Flag %}` is not an
    // `else`, and `elseif` is not a liquid tag at all.
    static Token BuildArgumentless(TokenKind kind, string tag, string source, string? expression, bool hasOtherContent)
    {
        var problem = BlockTagArguments.Problem(tag, expression);
        if (problem == null)
        {
            return new(kind, source, [], null, hasOtherContent);
        }

        return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent, Problem: problem);
    }

    static Token BuildForTag(string source, string? expression, bool hasOtherContent)
    {
        if (expression == null)
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        // Wrap the tag so Fluid yields a real ForStatement. We can then read the loop variable
        // (Identifier) directly off the AST instead of regex-parsing it.
        var liquid = $"{{% for {expression} %}}{{% endfor %}}";
        if (!parser.TryParse(liquid, out var template, out _))
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        var forStatement = ((FluidTemplate)template).Statements
            .OfType<ForStatement>()
            .FirstOrDefault();
        if (forStatement == null)
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        var references = IdentifierVisitor.Collect(forStatement.Source);
        var modifierReferences = IdentifierVisitor.Collect(forStatement.Limit);
        modifierReferences.AddRange(IdentifierVisitor.Collect(forStatement.Offset));
        var hasModifiers = forStatement.Limit != null ||
                           forStatement.Offset != null ||
                           forStatement.Reversed;
        return new(
            TokenKind.ForOpen,
            source,
            references,
            forStatement.Identifier,
            hasOtherContent,
            LoopModifierReferences: modifierReferences,
            HasLoopModifiers: hasModifiers,
            LoopSourceIsRange: forStatement.Source is not MemberExpression);
    }

    static Token BuildConditional(TokenKind kind, string source, string? expression, bool hasOtherContent)
    {
        if (expression == null)
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        var liquid = $"{{% if {expression} %}}{{% endif %}}";
        if (!parser.TryParse(liquid, out var template, out _))
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        var ifStatement = ((FluidTemplate)template).Statements
            .OfType<IfStatement>()
            .FirstOrDefault();
        if (ifStatement == null)
        {
            return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }

        var references = IdentifierVisitor.Collect(template);
        return new(kind, source, references, null, hasOtherContent);
    }
}
