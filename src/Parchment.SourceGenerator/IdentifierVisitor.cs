sealed class IdentifierVisitor :
    AstVisitor
{
    List<List<string>> paths = [];

    public static List<List<string>> Collect(IFluidTemplate template)
    {
        var visitor = new IdentifierVisitor();
        visitor.VisitTemplate(template);
        return visitor.paths;
    }

    public static List<List<string>> Collect(Expression? expression)
    {
        var visitor = new IdentifierVisitor();
        if (expression != null)
        {
            visitor.Visit(expression);
        }

        return visitor.paths;
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
