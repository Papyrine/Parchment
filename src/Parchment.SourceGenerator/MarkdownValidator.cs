/// <summary>
/// Walks a parsed markdown template's Fluid AST with loop-scope tracking, validating member
/// references against a <see cref="ModelShape"/>. Mirrors what <see cref="TokenScanner"/> +
/// <c>ParchmentTemplateGenerator.ValidateTokens</c> do for docx templates, but for markdown
/// the entire file is a single Fluid template (matching runtime Flow B), so we walk the AST
/// directly instead of scanning for token sites paragraph-by-paragraph.
///
/// Diagnostics emitted:
///   PARCH001 — MissingMember
///   PARCH002 — LoopSourceNotEnumerable
/// PARCH003 / PARCH005 / PARCH007 / PARCH008 / PARCH010 are docx-specific and not emitted
/// for markdown — the runtime markdown flow has no concept of paragraph boundaries, no
/// Excelsior dispatch, and no [Html]/[Markdown] structural replacement.
/// </summary>
static class MarkdownValidator
{
    public static void Validate(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        IFluidTemplate template)
    {
        var statements = ((FluidTemplate) template).Statements;
        var scope = new Dictionary<string, string>(StringComparer.Ordinal);
        WalkStatements(context, target, location, statements, scope);
    }

    static void WalkStatements(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        IReadOnlyList<Statement> statements,
        Dictionary<string, string> scope)
    {
        foreach (var statement in statements)
        {
            WalkStatement(context, target, location, statement, scope);
        }
    }

    static void WalkStatement(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        Statement statement,
        Dictionary<string, string> scope)
    {
        switch (statement)
        {
            case ForStatement forStatement:
                WalkFor(context, target, location, forStatement, scope);
                break;

            case IfStatement ifStatement:
                WalkIf(context, target, location, ifStatement, scope);
                break;

            case OutputStatement output:
                ValidateExpression(context, target, location, output.Expression, scope, "{{ ... }}");
                break;
        }
    }

    static void WalkFor(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        ForStatement forStatement,
        Dictionary<string, string> scope)
    {
        var sourceText = $"{{% for {forStatement.Identifier} in ... %}}";

        // Catch any member-access references inside the source expression first (e.g. for a
        // filtered source `Customer.Lines | where: ...` we still want PARCH001 on Customer.Lines).
        ValidateExpression(context, target, location, forStatement.Source, scope, sourceText);

        string? elementFqn = null;
        var sourcePath = TryGetMemberPath(forStatement.Source);
        if (sourcePath != null)
        {
            var sourceFqn = ShapeResolver.Resolve(target.Shape, sourcePath, scope);
            if (sourceFqn != null)
            {
                elementFqn = ShapeResolver.GetElementType(target.Shape, sourceFqn);
                if (elementFqn == null)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.LoopSourceNotEnumerable,
                            location,
                            target.TemplatePath,
                            sourceText));
                }
            }
        }

        var loopVariable = forStatement.Identifier;
        var hadPrior = scope.TryGetValue(loopVariable, out var prior);
        // When the source didn't resolve to an enumerable, bind the loop variable to the root
        // type. That's wrong but it minimises cascade noise — accesses on the loop variable
        // just resolve against the root model instead of generating a wave of PARCH001s on top
        // of the upstream PARCH001/PARCH002 that already reported the real problem.
        scope[loopVariable] = elementFqn ?? target.Shape.RootTypeFullyQualifiedName;

        WalkStatements(context, target, location, forStatement.Statements, scope);
        if (forStatement.Else != null)
        {
            WalkStatements(context, target, location, forStatement.Else.Statements, scope);
        }

        if (hadPrior)
        {
            scope[loopVariable] = prior!;
        }
        else
        {
            scope.Remove(loopVariable);
        }
    }

    static void WalkIf(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        IfStatement ifStatement,
        Dictionary<string, string> scope)
    {
        ValidateExpression(context, target, location, ifStatement.Condition, scope, "{% if ... %}");
        WalkStatements(context, target, location, ifStatement.Statements, scope);

        foreach (var branch in ifStatement.ElseIfs)
        {
            ValidateExpression(context, target, location, branch.Condition, scope, "{% elsif ... %}");
            WalkStatements(context, target, location, branch.Statements, scope);
        }

        if (ifStatement.Else != null)
        {
            WalkStatements(context, target, location, ifStatement.Else.Statements, scope);
        }
    }

    static void ValidateExpression(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        Expression expression,
        Dictionary<string, string> scope,
        string sourceForDiagnostic)
    {
        var paths = ExpressionPathCollector.Collect(expression);
        foreach (var path in paths)
        {
            var resolved = ShapeResolver.Resolve(target.Shape, path, scope);
            if (resolved != null)
            {
                continue;
            }

            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.MissingMember,
                    location,
                    target.TemplatePath,
                    sourceForDiagnostic,
                    string.Join('.', path),
                    target.ModelDisplayName));
        }
    }

    static List<string>? TryGetMemberPath(Expression expression)
    {
        if (expression is not MemberExpression member)
        {
            return null;
        }

        // Use the shared extractor so indexer-with-string-literal (`Customer['Lines']`) resolves
        // the same as dotted access (`Customer.Lines`). Without this, a loop source written with
        // bracket notation would fail to resolve its element type and the body walk would bind
        // the loop variable to the root model, producing false PARCH001s on every body access.
        var segments = new List<string>(member.Segments.Count);
        foreach (var segment in member.Segments)
        {
            var name = SegmentNames.TryGetStaticName(segment);
            if (name == null)
            {
                return null;
            }

            segments.Add(name);
        }

        return segments.Count == 0 ? null : segments;
    }

    sealed class ExpressionPathCollector :
        AstVisitor
    {
        List<List<string>> paths = [];

        public static List<List<string>> Collect(Expression expression)
        {
            var collector = new ExpressionPathCollector();
            collector.Visit(expression);
            return collector.paths;
        }

        protected override Expression VisitMemberExpression(MemberExpression memberExpression)
        {
            var segments = new List<string>(memberExpression.Segments.Count);
            foreach (var segment in memberExpression.Segments)
            {
                var name = SegmentNames.TryGetStaticName(segment);
                if (name == null)
                {
                    break;
                }

                segments.Add(name);
            }

            if (segments.Count > 0)
            {
                paths.Add(segments);
            }

            return base.VisitMemberExpression(memberExpression);
        }
    }
}
