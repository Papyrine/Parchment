sealed record IdentifierPath(IReadOnlyList<string> Segments)
{
    public string Root => Segments[0];

    // Lazily memoized — IdentifierPath is registration-time data, queried per render via
    // TryResolveExcelsior / TryResolveFormatted / TryResolveStringList. Computing the join
    // once avoids ~3 string.Join allocations per token per render. Backing field is a private
    // mutable cache rather than a record auto-property to keep it out of synthesized equality.
    public string Dotted => field ??= string.Join('.', Segments);

    public override string ToString() => Dotted;
}
