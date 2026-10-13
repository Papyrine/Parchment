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
///   PARCH003 — UnsupportedBlockTag, for the partial tags `render` and `include` only
///   PARCH024 — FormatMarkerThroughLoopVariable
///   PARCH037 — FilterAfterTokenFilter
/// PARCH005 / PARCH010 are docx-specific and not emitted for markdown — the runtime
/// markdown flow has no concept of paragraph boundaries and no [Html]/[Markdown] structural
/// replacement.
///
/// PARCH007 / PARCH008 do apply to markdown, but not from here: an Excelsior token is rewritten
/// to a marker before the source is parsed, so what disqualifies one is a property of the line it
/// sits on rather than of this AST. ParchmentTemplateGenerator.ValidateMarkdownExcelsiorTokens
/// reads them off the source instead.
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
    readonly string templatePath;
    readonly Location location;

    // Loop variables, bound to the fully qualified name of the element type being iterated.
    readonly Dictionary<string, string> scope = new(StringComparer.Ordinal);

    // Identifiers that exist but whose type is not knowable statically: assign/capture targets,
    // `forloop`, and loop variables over a source that is itself untyped. Members off these are not
    // validated, but the root is not reported as missing either.
    readonly HashSet<string> untyped = new(StringComparer.Ordinal);

    MarkdownValidator(SourceProductionContext context, TargetInfo target, string templatePath, Location location)
    {
        this.context = context;
        this.target = target;
        this.templatePath = templatePath;
        this.location = location;
    }

    public static void Validate(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        Location location,
        IFluidTemplate template)
    {
        var validator = new MarkdownValidator(context, target, templatePath, location);
        validator.ReportPartialTags(template);
        validator.WalkStatements(((FluidTemplate) template).Statements);
    }

    // `{% render %}` and `{% include %}` parse, and then load another template through a file
    // provider nothing configures, so the runtime refuses them at registration. Found with a
    // visitor of their own rather than in WalkStatement, which descends into the statements it
    // validates references in and no others: a partial is unsupported wherever it sits.
    void ReportPartialTags(IFluidTemplate template)
    {
        foreach (var tag in PartialTagCollector.Collect(template))
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.UnsupportedBlockTag,
                    location,
                    templatePath,
                    $"{{% {tag} ... %}}",
                    Diagnostics.MarkdownPartialTag,
                    string.Empty));
        }
    }

    sealed class PartialTagCollector :
        AstVisitor
    {
        List<string> tags = [];

        public static List<string> Collect(IFluidTemplate template)
        {
            var collector = new PartialTagCollector();
            collector.VisitTemplate(template);
            return collector.tags;
        }

        protected override Statement VisitIncludeStatement(IncludeStatement includeStatement)
        {
            tags.Add("include");
            return includeStatement;
        }

        protected override Statement VisitRenderStatement(RenderStatement renderStatement)
        {
            tags.Add("render");
            return renderStatement;
        }
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

            case UnlessStatement unlessStatement:
                WalkUnless(unlessStatement);
                break;

            case CaseStatement caseStatement:
                WalkCase(caseStatement);
                break;

            case CycleStatement cycleStatement:
                WalkCycle(cycleStatement);
                break;

            case TableRowStatement tableRowStatement:
                WalkTableRow(tableRowStatement);
                break;

            // `{% liquid %}` and `{% ifchanged %}` hold statements and nothing else. Everything
            // left over has no expression and no body to walk: text, `raw`, `comment`, `break`,
            // `continue`, and `increment` / `decrement`, whose counter is not a variable a token
            // can read.
            case LiquidStatement or IfChangedStatement:
                WalkStatements(((TagStatement) statement).Statements);
                break;
        }
    }

    void WalkFor(ForStatement forStatement)
    {
        var sourceText = $"{{% for {forStatement.Identifier} in ... %}}";

        // Catch any member-access references inside the source expression first (e.g. for a
        // filtered source `Customer.Lines | where: ...` we still want PARCH001 on Customer.Lines).
        ValidateExpression(forStatement.Source, sourceText);

        // `limit:` and `offset:` are evaluated before the loop variable is bound, so they resolve
        // against the enclosing scope. `offset: continue` is liquid's own keyword, which Fluid
        // parses as a member path.
        if (forStatement.Limit != null)
        {
            ValidateExpression(forStatement.Limit, sourceText);
        }

        if (forStatement.Offset != null &&
            !forStatement.OffsetIsContinue)
        {
            ValidateExpression(forStatement.Offset, sourceText);
        }

        string? elementFqn = null;
        var sourcePath = TryGetMemberPath(forStatement.Source);
        // Iterating an untyped identifier — the target of an earlier assign, say. Nothing about its
        // shape is known, so neither the source nor the loop variable can be checked.
        var sourceIsUntyped = sourcePath != null && untyped.Contains(sourcePath[0]);
        if (sourcePath != null &&
            !sourceIsUntyped &&
            ShapeResolver.TryResolve(target.Shape, sourcePath, scope, out var sourceFqn))
        {
            if (sourceFqn == ShapeResolver.Untyped)
            {
                // Reached through an index that says nothing about what it yields, a dictionary
                // key for one. Unknown in the same way, and not a source to report.
                sourceIsUntyped = true;
            }
            else if (!ShapeResolver.TryGetElementType(target.Shape, sourceFqn, out elementFqn))
            {
                context.ReportDiagnostic(
                    Diagnostic.Create(
                        Diagnostics.LoopSourceNotEnumerable,
                        location,
                        templatePath,
                        sourceText));
            }
        }

        var loopVariable = forStatement.Identifier;
        var hadPrior = scope.TryGetValue(loopVariable, out var prior);
        var hadUntyped = untyped.Contains(loopVariable);
        // `forloop` is introduced by liquid for the duration of the body.
        var addedForLoop = untyped.Add("forloop");

        if (sourceIsUntyped ||
            sourcePath == null)
        {
            // Nothing is known about what is being iterated: an untyped identifier, or a source
            // with no member path at all such as the range in `{% for i in (1..5) %}`. Neither
            // reported a diagnostic above, so falling through to the root-model binding below would
            // report a PARCH001 on every member reached through the loop variable, none of them
            // real.
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

        // The else branch renders when there was nothing to iterate, so neither the loop variable
        // nor `forloop` is bound in it: it is walked in the enclosing scope.
        if (forStatement.Else != null)
        {
            WalkStatements(forStatement.Else.Statements);
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

    void WalkUnless(UnlessStatement unlessStatement)
    {
        ValidateExpression(unlessStatement.Condition, "{% unless ... %}");
        WalkStatements(unlessStatement.Statements);

        if (unlessStatement.Else != null)
        {
            WalkStatements(unlessStatement.Else.Statements);
        }
    }

    void WalkCase(CaseStatement caseStatement)
    {
        ValidateExpression(caseStatement.Expression, "{% case ... %}");

        foreach (var when in caseStatement.Whens)
        {
            foreach (var option in when.Options)
            {
                ValidateExpression(option, "{% when ... %}");
            }

            WalkStatements(when.Statements);
        }

        if (caseStatement.Else != null)
        {
            WalkStatements(caseStatement.Else.Statements);
        }
    }

    void WalkCycle(CycleStatement cycleStatement)
    {
        if (cycleStatement.Group != null)
        {
            ValidateExpression(cycleStatement.Group, "{% cycle ... %}");
        }

        foreach (var value in cycleStatement.Values)
        {
            ValidateExpression(value, "{% cycle ... %}");
        }
    }

    // A loop like `for`, with three differences: no else branch, `tablerowloop` in place of
    // `forloop`, and arguments that are validated.
    void WalkTableRow(TableRowStatement tableRowStatement)
    {
        var loopVariable = tableRowStatement.Identifier;
        var sourceText = $"{{% tablerow {loopVariable} in ... %}}";

        // The source and the arguments are evaluated before the first cell, so they validate
        // against the enclosing scope: `limit: row.Count` does not see the row.
        ValidateExpression(tableRowStatement.Source, sourceText);
        ValidateArgument(tableRowStatement.Limit, sourceText);
        ValidateArgument(tableRowStatement.Offset, sourceText);
        ValidateArgument(tableRowStatement.Cols, sourceText);

        string? elementFqn = null;
        var sourcePath = TryGetMemberPath(tableRowStatement.Source);
        var sourceIsUntyped = sourcePath != null && untyped.Contains(sourcePath[0]);
        if (sourcePath != null &&
            !sourceIsUntyped &&
            ShapeResolver.TryResolve(target.Shape, sourcePath, scope, out var sourceFqn) &&
            !ShapeResolver.TryGetElementType(target.Shape, sourceFqn, out elementFqn))
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.LoopSourceNotEnumerable,
                    location,
                    templatePath,
                    sourceText));
        }

        var hadPrior = scope.TryGetValue(loopVariable, out var prior);
        var hadUntyped = untyped.Contains(loopVariable);
        var addedTableRowLoop = untyped.Add("tablerowloop");

        if (sourceIsUntyped ||
            sourcePath == null)
        {
            scope.Remove(loopVariable);
            untyped.Add(loopVariable);
        }
        else
        {
            // The root type when the source did not resolve, for the reason WalkFor gives.
            untyped.Remove(loopVariable);
            scope[loopVariable] = elementFqn ?? target.Shape.RootTypeFullyQualifiedName;
        }

        WalkStatements(tableRowStatement.Statements);

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

        if (addedTableRowLoop)
        {
            untyped.Remove("tablerowloop");
        }
    }

    void ValidateArgument(Expression? argument, string sourceForDiagnostic)
    {
        if (argument != null)
        {
            ValidateExpression(argument, sourceForDiagnostic);
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
        ValidateFilterOrder(expression);

        var paths = ExpressionPathCollector.Collect(expression);
        foreach (var path in paths)
        {
            if (untyped.Contains(path[0]))
            {
                continue;
            }

            if (ShapeResolver.TryResolve(target.Shape, path, scope, out _))
            {
                WarnIfFormatMarkerIsUnreachable(path, sourceForDiagnostic);
                continue;
            }

            context.ReportDiagnostic(
                Diagnostics.UnknownMember(target, templatePath, location, sourceForDiagnostic, path, scope));
        }
    }

    // `html` is the one filter that produces a token in this flow: it parks its value and writes a
    // placeholder, which a filter after it would be handed in place of the value. Lockstep with
    // MarkdownReferenceValidator.VisitFilterExpression. A chain is only ever the whole of an output
    // or an assign value, so the expressions handed to ValidateExpression are everywhere one can be.
    void ValidateFilterOrder(Expression expression)
    {
        if (!TokenFilterOrder.TryFind(expression, markdownFlow: true, out var tokenFilter, out var following))
        {
            return;
        }

        context.ReportDiagnostic(
            Diagnostic.Create(
                Diagnostics.FilterAfterTokenFilter,
                location,
                templatePath,
                TokenFilterOrder.Describe(expression),
                TokenFilterOrder.Message(tokenFilter, following)));
    }

    // A format marker on a member reached through a loop variable cannot be honoured: the rewrite
    // that applies one matches a token's path against the map built from the model root, and a
    // loop variable is not addressable from there. The token renders as escaped text with its
    // markup visible, which is the sort of thing worth being told about rather than discovering in
    // a rendered document.
    void WarnIfFormatMarkerIsUnreachable(IReadOnlyList<string> path, string sourceForDiagnostic)
    {
        if (!scope.ContainsKey(path[0]) ||
            !ShapeResolver.TryResolveMember(target.Shape, path, scope, out var member) ||
            member is
            {
                IsHtml: false,
                IsMarkdown: false
            })
        {
            return;
        }

        context.ReportDiagnostic(
            Diagnostic.Create(
                Diagnostics.FormatMarkerThroughLoopVariable,
                location,
                templatePath,
                sourceForDiagnostic,
                member.Name));
    }

    // Returns null for a source with no static member path — a range. WalkFor treats that as
    // untyped rather than binding to the root model. An indexer whose key is not a string literal
    // stays in the path as SegmentNames.Index, and whether the loop variable can be typed through
    // it is ShapeResolver.TryResolve's to say.
    // A filtered source needs no handling here: Fluid rejects `{% for x in Src | filter %}` as an
    // invalid for tag, so it never reaches validation on either the runtime or the generator path.
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

        if (segments.Count == 0)
        {
            return null;
        }

        return segments;
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
