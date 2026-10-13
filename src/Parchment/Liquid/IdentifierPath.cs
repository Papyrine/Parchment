sealed record IdentifierPath(IReadOnlyList<string> Segments)
{
    /// <summary>
    /// The segment an indexer leaves in a path unless it is a string literal: <c>Tags[0]</c>,
    /// <c>Lines[i]</c>. What follows it is read from an element, so the validators step into the
    /// element type there, and a path holding one is never the member it indexes into — no key of
    /// the root-keyed attribute maps has a bracket in it, so <see cref="Dotted"/> matches none.
    /// </summary>
    public const string Index = "[]";

    public string Root => Segments[0];

    // Lazily memoized — IdentifierPath is registration-time data, queried per render via
    // TryResolveExcelsior / TryResolveFormatted / TryResolveStringList. Computing the join
    // once avoids ~3 string.Join allocations per token per render. Backing field is a private
    // mutable cache rather than a record auto-property to keep it out of synthesized equality.
    // An index is written against the segment it applies to: `Rows[].Item`.
    public string Dotted => field ??= string.Join('.', Segments).Replace("." + Index, Index);

    public override string ToString() => Dotted;
}
