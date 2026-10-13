// Shared between IdentifierVisitor and MarkdownValidator.ExpressionPathCollector. Indexer with a
// string literal (`Customer['Name']`) is semantically identical to dotted access (`Customer.Name`)
// — Fluid resolves both the same way at render time. Extracting the literal here closes the
// validation gap where `Customer['NoSuchMember']` would silently render empty because the visitor
// previously only walked `IdentifierSegment`s. Any other indexer (numeric, variable, expression)
// picks an element at render. It stays in the path as Index rather than ending it: cut off there,
// `Tags[0]` reads as `Tags` and the token is taken for the member it indexes into. Mirrors the
// runtime's IdentifierVisitor.TryGetStaticName and IdentifierPath.Index.
static class SegmentNames
{
    // No member is spelled this way, so a path holding one resolves to no MemberEntry: the rules
    // for an [ExcelsiorTable], [Html] / [Markdown] or [EditableField] member's own token never
    // apply to a token that indexes into it. ShapeResolver.TryResolve steps into the element type.
    public const string Index = "[]";

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
            IndexerSegment => Index,
            _ => null
        };

    // A path as a diagnostic shows it, the index written against the segment it applies to:
    // `Rows[].Item`.
    public static string Join(IReadOnlyList<string> path) =>
        string.Join(".", path).Replace("." + Index, Index);
}
