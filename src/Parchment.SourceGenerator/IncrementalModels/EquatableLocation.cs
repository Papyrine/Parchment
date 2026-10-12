readonly record struct EquatableLocation(
    string? FilePath,
    TextSpan TextSpan,
    LinePositionSpan LineSpan)
{
    public static EquatableLocation From(Location location)
    {
        if (location == Location.None)
        {
            return new(null, default, default);
        }

        var mapped = location.GetLineSpan();
        return new(mapped.Path, location.SourceSpan, mapped.Span);
    }

    /// <summary>
    /// Rebuilds the location inside <paramref name="tree"/>, the tree it was taken from.
    /// </summary>
    /// <remarks>
    /// A location built from a path alone is an external-file location: it has no syntax tree, so
    /// the compiler cannot match it against a <c>#pragma warning disable</c>, a
    /// <c>[SuppressMessage]</c> or a per-file <c>.editorconfig</c> severity. That form is kept
    /// only as the fallback for a missing tree.
    /// </remarks>
    public Location ToLocation(SyntaxTree? tree)
    {
        if (FilePath is null)
        {
            return Location.None;
        }

        if (tree != null &&
            TextSpan.End <= tree.Length)
        {
            return Location.Create(tree, TextSpan);
        }

        return Location.Create(FilePath, TextSpan, LineSpan);
    }
}
