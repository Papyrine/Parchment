/// <summary>
/// Walks a Fluid AST collecting the root identifiers of every member expression it sees.
/// Used by ModelValidator to check that referenced model members actually exist.
/// </summary>
class IdentifierVisitor :
    AstVisitor
{
    readonly List<IdentifierPath> paths = [];

    public static IReadOnlyList<IdentifierPath> Collect(IFluidTemplate template)
    {
        var visitor = new IdentifierVisitor();
        visitor.VisitTemplate(template);
        return visitor.paths;
    }

    public static IReadOnlyList<IdentifierPath> Collect(Expression expression)
    {
        var visitor = new IdentifierVisitor();
        visitor.Visit(expression);
        return visitor.paths;
    }

    protected override Expression VisitMemberExpression(MemberExpression expression)
    {
        var segments = new List<string>(expression.Segments.Count);
        foreach (var segment in expression.Segments)
        {
            if (segment is IdentifierSegment identifier)
            {
                segments.Add(identifier.Identifier);
            }
            else
            {
                break;
            }
        }

        if (segments.Count > 0)
        {
            paths.Add(new(segments));
        }

        return base.VisitMemberExpression(expression);
    }
}
