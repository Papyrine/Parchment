sealed record LoopNode(
    string OpenAnchorName,
    string CloseAnchorName,
    string LoopVariable,
    Expression LoopSource,
    // First dotted identifier of LoopSource (e.g. `Customer.Lines`), or null when the source is not
    // a plain member path. Precomputed at tree-build time so the render path can probe the editable
    // collection map without re-walking the AST via IdentifierVisitor.Collect on every loop.
    string? SourceReference,
    IReadOnlyList<RangeNode> Body) :
    RangeNode;
