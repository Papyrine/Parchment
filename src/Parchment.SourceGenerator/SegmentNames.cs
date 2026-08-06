// Shared between IdentifierVisitor and MarkdownValidator.ExpressionPathCollector. Indexer with a
// string literal (`Customer['Name']`) is semantically identical to dotted access (`Customer.Name`)
// — Fluid resolves both the same way at render time. Extracting the literal here closes the
// validation gap where `Customer['NoSuchMember']` would silently render empty because the visitor
// previously only walked `IdentifierSegment`s. Non-literal indexers (variable, numeric, expression)
// still terminate the path — those are either dictionary/array access (legitimate, unvalidatable
// statically) or typos we can't catch.
static class SegmentNames
{
    public static string? TryGetStaticName(MemberSegment segment) =>
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
            _ => null
        };
}
