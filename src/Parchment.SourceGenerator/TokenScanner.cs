public static class TokenScanner
{
    static readonly FluidParser parser = new();

    // containers runs parallel to paragraphs: the id of the element each paragraph sits in, copied
    // onto its tokens. Omitted, every token reports container 0.
    // parts runs parallel the same way: the index of the part each paragraph came from.
    public static List<Token> Scan(
        IEnumerable<string> paragraphs,
        IReadOnlyList<int>? containers = null,
        IReadOnlyList<int>? parts = null)
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
            // The runtime rejects a block tag that shares its paragraph with anything: text, a
            // substitution, or another block tag. Only the first is content outside the sites.
            var blockHasOtherContent = hasOtherContent || sites.Count > 1;
            var container = 0;
            if (containers != null)
            {
                container = containers[index];
            }

            var part = 0;
            if (parts != null)
            {
                part = parts[index];
            }

            foreach (var site in sites)
            {
                var source = paragraph.Substring(site.Offset, site.Length);
                if (site.Kind == TokenSiteKind.Substitution)
                {
                    result.Add(ParseSubstitution(source, hasOtherContent) with { Container = container, Part = part });
                    continue;
                }

                result.Add(ParseBlockTag(source, blockHasOtherContent) with { Container = container, Part = part });
            }
        }

        return result;
    }

    static Token ParseSubstitution(string source, bool hasOtherContent)
    {
        if (!parser.TryParse(source, out var template, out var error))
        {
            return new(TokenKind.Substitution, source, [], null, hasOtherContent, ParseError: error);
        }

        var references = IdentifierVisitor.Collect(template);
        var isPlain = IsPlainMemberAccess(template);
        return new(TokenKind.Substitution, source, references, null, hasOtherContent, isPlain);
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
                return new(TokenKind.ForClose, source, [], null, hasOtherContent);

            case "if":
                return BuildConditional(TokenKind.IfOpen, source, expression, hasOtherContent);

            case "elsif":
            case "elseif":
                return BuildConditional(TokenKind.ElsIf, source, expression, hasOtherContent);

            case "else":
                return new(TokenKind.Else, source, [], null, hasOtherContent);

            case "endif":
                return new(TokenKind.IfClose, source, [], null, hasOtherContent);

            default:
                return new(TokenKind.UnknownBlock, source, [], null, hasOtherContent);
        }
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
