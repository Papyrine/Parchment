/// <summary>
/// Validates every member reference in a markdown template against the model, binding loop
/// variables to their element type so members reached through a loop are checked too.
/// </summary>
/// <remarks>
/// This replaces a raw-text scan for <c>{% for &lt;name&gt; in </c>, which was wrong in both
/// directions. It missed loop variables under leading whitespace control (<c>{%- for</c>), failing
/// registration for a valid template; and when it did match, the whole subtree was skipped, so a
/// typo'd member on a loop variable threw nothing at registration or render — the value resolved to
/// nil and the content silently vanished. Walking the AST with a scope mirrors what
/// <see cref="ReferenceValidator"/> already does for the docx flow and what the source generator
/// does for markdown at compile time.
/// </remarks>
class MarkdownReferenceValidator :
    AstVisitor
{
    readonly Type modelType;
    readonly string templateName;

    // Loop variables, bound to the element type of whatever is being iterated.
    readonly Dictionary<string, Type> scope = new(StringComparer.Ordinal);

    // Identifiers that exist but whose type is not knowable statically: assign/capture targets,
    // `forloop`, and loop variables over a source that is not a resolvable enumerable (a range, for
    // instance). Members off these are not validated, but the root is not reported as unknown.
    readonly HashSet<string> untyped = new(StringComparer.Ordinal);

    MarkdownReferenceValidator(Type modelType, string templateName)
    {
        this.modelType = modelType;
        this.templateName = templateName;
    }

    public static void Validate(IFluidTemplate template, Type modelType, string templateName)
    {
        var validator = new MarkdownReferenceValidator(modelType, templateName);
        validator.VisitTemplate(template);
    }

    protected override Statement VisitForStatement(ForStatement forStatement)
    {
        // The source is evaluated outside the loop, so it validates against the enclosing scope
        // before the loop variable is introduced.
        Visit(forStatement.Source);

        var elementType = ResolveElementType(forStatement.Source);

        var hadScope = scope.TryGetValue(forStatement.Identifier, out var previousScope);
        var hadUntyped = untyped.Contains(forStatement.Identifier);
        var addedForLoop = untyped.Add("forloop");

        if (elementType == null)
        {
            // The source is not a resolvable enumerable (a range, or a value off an assign). The
            // loop variable still exists, so accept it without checking its members.
            scope.Remove(forStatement.Identifier);
            untyped.Add(forStatement.Identifier);
        }
        else
        {
            untyped.Remove(forStatement.Identifier);
            scope[forStatement.Identifier] = elementType;
        }

        foreach (var statement in forStatement.Statements)
        {
            Visit(statement);
        }

        // The else branch runs when the source is empty. Overriding VisitForStatement means base is
        // never called, so without this the whole branch goes unvisited and a typo inside it throws
        // nothing — the same silent skip this class was written to remove.
        if (forStatement.Else != null)
        {
            foreach (var statement in forStatement.Else.Statements)
            {
                Visit(statement);
            }
        }

        // Restore, so a name bound by one loop does not leak into a sibling or an enclosing scope.
        scope.Remove(forStatement.Identifier);
        untyped.Remove(forStatement.Identifier);
        if (hadScope)
        {
            scope[forStatement.Identifier] = previousScope!;
        }

        if (hadUntyped)
        {
            untyped.Add(forStatement.Identifier);
        }

        if (addedForLoop)
        {
            untyped.Remove("forloop");
        }

        return forStatement;
    }

    // `assign` and `capture` introduce a variable Fluid holds as a plain string identifier, not a
    // MemberExpression, so the tag itself never reached the validator — only the later use did, as
    // an unresolvable root.
    protected override Statement VisitAssignStatement(AssignStatement assignStatement)
    {
        Visit(assignStatement.Value);
        untyped.Add(assignStatement.Identifier);
        return assignStatement;
    }

    protected override Statement VisitCaptureStatement(CaptureStatement captureStatement)
    {
        foreach (var statement in captureStatement.Statements)
        {
            Visit(statement);
        }

        untyped.Add(captureStatement.Identifier);
        return captureStatement;
    }

    protected override Expression VisitMemberExpression(MemberExpression expression)
    {
        var segments = new List<string>(expression.Segments.Count);
        foreach (var segment in expression.Segments)
        {
            var name = IdentifierVisitor.TryGetStaticName(segment);
            if (name == null)
            {
                break;
            }

            segments.Add(name);
        }

        if (segments.Count > 0 &&
            !untyped.Contains(segments[0]))
        {
            ModelValidator.Validate(modelType, new(segments), scope, templateName, null, null);
        }

        return base.VisitMemberExpression(expression);
    }

    Type? ResolveElementType(Expression source)
    {
        var refs = IdentifierVisitor.Collect(source);
        if (refs.Count == 0)
        {
            return null;
        }

        var iterableType = ResolvePathType(refs[0]);
        if (iterableType == null)
        {
            return null;
        }

        return ModelValidator.TryResolveElementType(iterableType);
    }

    Type? ResolvePathType(IdentifierPath path)
    {
        if (untyped.Contains(path.Root))
        {
            return null;
        }

        var current = scope.TryGetValue(path.Root, out var scoped)
            ? scoped
            : ModelValidator.ResolveMember(modelType, path.Root);
        if (current == null)
        {
            return null;
        }

        for (var i = 1; i < path.Segments.Count; i++)
        {
            var next = ModelValidator.ResolveMember(current, path.Segments[i]);
            if (next == null)
            {
                return null;
            }

            current = next;
        }

        return current;
    }
}
