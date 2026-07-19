/// <summary>
/// Walks a parsed markdown template's Fluid AST with loop-scope tracking, validating member
/// references against a <see cref="ModelShape"/>. Mirrors what <see cref="TokenScanner"/> +
/// <c>ParchmentTemplateGenerator.ValidateTokens</c> do for docx templates, but for markdown
/// the entire file is a single Fluid template (matching runtime Flow B), so the AST is walked
/// directly instead of scanning for token sites paragraph-by-paragraph.
///
/// Diagnostics emitted:
///   PARCH001 — MissingMember
///   PARCH002 — LoopSourceNotEnumerable
/// PARCH003 / PARCH005 / PARCH007 / PARCH008 / PARCH010 are docx-specific and not emitted
/// for markdown — the runtime markdown flow has no concept of paragraph boundaries, no
/// Excelsior dispatch, and no [Html]/[Markdown] structural replacement.
/// </summary>
/// <remarks>
/// Kept deliberately in step with <c>Parchment.Liquid.MarkdownReferenceValidator</c>, its runtime
/// counterpart. The two validate the same templates against the same model, so whatever one accepts
/// the other has to accept. Divergence here is the worse direction of the two: PARCH001 is an
/// error, so a template that registers at runtime but is rejected here fails the build outright.
/// The scope/untyped handling below is structured to match that class method for method.
/// </remarks>
class MarkdownValidator
{
    readonly SourceProductionContext context;
    readonly TargetInfo target;
    readonly Location location;

    // Loop variables, bound to the fully qualified name of the element type being iterated.
    readonly Dictionary<string, string> scope = new(StringComparer.Ordinal);

    // Identifiers that exist but whose type is not knowable statically: assign/capture targets,
    // `forloop`, and loop variables over a source that is itself untyped. Members off these are not
    // validated, but the root is not reported as missing either.
    readonly HashSet<string> untyped = new(StringComparer.Ordinal);

    MarkdownValidator(SourceProductionContext context, TargetInfo target, Location location)
    {
        this.context = context;
        this.target = target;
        this.location = location;
    }

    public static void Validate(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        IFluidTemplate template)
    {
        var validator = new MarkdownValidator(context, target, location);
        validator.WalkStatements(((FluidTemplate) template).Statements);
    }

    void WalkStatements(IReadOnlyList<Statement> statements)
    {
        foreach (var statement in statements)
        {
            WalkStatement(statement);
        }
    }

    void WalkStatement(Statement statement)
    {
        switch (statement)
        {
            case ForStatement forStatement:
                WalkFor(forStatement);
                break;

            case IfStatement ifStatement:
                WalkIf(ifStatement);
                break;

            case AssignStatement assignStatement:
                WalkAssign(assignStatement);
                break;

            case CaptureStatement captureStatement:
                WalkCapture(captureStatement);
                break;

            case OutputStatement output:
                ValidateExpression(output.Expression, "{{ ... }}");
                break;
        }
    }

    void WalkFor(ForStatement forStatement)
    {
        var sourceText = $"{{% for {forStatement.Identifier} in ... %}}";

        // Catch any member-access references inside the source expression first (e.g. for a
        // filtered source `Customer.Lines | where: ...` we still want PARCH001 on Customer.Lines).
        ValidateExpression(forStatement.Source, sourceText);

        string? elementFqn = null;
        var sourcePath = TryGetMemberPath(forStatement.Source);
        // Iterating an untyped identifier — the target of an earlier assign, say. Nothing about its
        // shape is known, so neither the source nor the loop variable can be checked.
        var sourceIsUntyped = sourcePath != null && untyped.Contains(sourcePath[0]);
        if (sourcePath != null &&
            !sourceIsUntyped)
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
        var hadUntyped = untyped.Contains(loopVariable);
        // `forloop` is introduced by liquid for the duration of the body.
        var addedForLoop = untyped.Add("forloop");

        if (sourceIsUntyped)
        {
            // Falling through to the root-model binding below would report a PARCH001 on every
            // member reached through the loop variable, none of them real.
            scope.Remove(loopVariable);
            untyped.Add(loopVariable);
        }
        else
        {
            // When the source didn't resolve to an enumerable, bind the loop variable to the root
            // type. That's wrong but it minimises cascade noise — accesses on the loop variable
            // resolve against the root model instead of generating a wave of PARCH001s on top
            // of the upstream PARCH001/PARCH002 that already reported the real problem.
            untyped.Remove(loopVariable);
            scope[loopVariable] = elementFqn ?? target.Shape.RootTypeFullyQualifiedName;
        }

        WalkStatements(forStatement.Statements);
        if (forStatement.Else != null)
        {
            WalkStatements(forStatement.Else.Statements);
        }

        // Restore, so a name bound by one loop does not leak into a sibling or an enclosing scope.
        scope.Remove(loopVariable);
        untyped.Remove(loopVariable);
        if (hadPrior)
        {
            scope[loopVariable] = prior!;
        }

        if (hadUntyped)
        {
            untyped.Add(loopVariable);
        }

        if (addedForLoop)
        {
            untyped.Remove("forloop");
        }
    }

    void WalkIf(IfStatement ifStatement)
    {
        ValidateExpression(ifStatement.Condition, "{% if ... %}");
        WalkStatements(ifStatement.Statements);

        foreach (var branch in ifStatement.ElseIfs)
        {
            ValidateExpression(branch.Condition, "{% elsif ... %}");
            WalkStatements(branch.Statements);
        }

        if (ifStatement.Else != null)
        {
            WalkStatements(ifStatement.Else.Statements);
        }
    }

    // `assign` and `capture` introduce a variable Fluid holds as a plain string identifier rather
    // than a MemberExpression, so the tag itself never reaches ValidateExpression — only the later
    // use does, as a root with nothing to resolve against.
    void WalkAssign(AssignStatement assignStatement)
    {
        // The value is a normal expression and stays checked: `{% assign total = NoSuchThing %}`
        // is still a PARCH001.
        ValidateExpression(assignStatement.Value, "{% assign ... %}");
        untyped.Add(assignStatement.Identifier);
    }

    void WalkCapture(CaptureStatement captureStatement)
    {
        WalkStatements(captureStatement.Statements);
        untyped.Add(captureStatement.Identifier);
    }

    void ValidateExpression(Expression expression, string sourceForDiagnostic)
    {
        var paths = ExpressionPathCollector.Collect(expression);
        foreach (var path in paths)
        {
            if (untyped.Contains(path[0]))
            {
                continue;
            }

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
