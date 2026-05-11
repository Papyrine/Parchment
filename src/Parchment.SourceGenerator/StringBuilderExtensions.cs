static class StringBuilderExtensions
{
    public static StringBuilder Indent(this StringBuilder builder, int depth) =>
        builder.Append(' ', depth * 2);
}
