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
            var name = TryGetStaticName(segment);
            if (name == null)
            {
                break;
            }

            segments.Add(name);
        }

        if (segments.Count > 0)
        {
            paths.Add(new(segments));
        }

        return base.VisitMemberExpression(expression);
    }

    // Indexer with a string literal (`Customer['Name']`) is semantically identical to dotted access
    // (`Customer.Name`) — Fluid resolves both the same way at render time. Extracting the literal
    // here closes the validation gap where `Customer['NoSuchMember']` would silently render empty
    // because IdentifierVisitor only walked `IdentifierSegment`s and stopped at the indexer.
    // Any other indexer (numeric, variable, expression) picks an element at render. It stays in
    // the path as IdentifierPath.Index rather than ending it: cut off there, `Tags[0]` reads as
    // `Tags` and the token is taken for the member it indexes into.
    internal static string? TryGetStaticName(MemberSegment segment) =>
        segment switch
        {
            IdentifierSegment id => id.Identifier,
            IndexerSegment
            {
                Expression: LiteralExpression
                {
                    Value:
                    {
                        Type: FluidValues.String
                    } value
                }
            } => value.ToStringValue(),
            IndexerSegment => IdentifierPath.Index,
            _ => null
        };
}
