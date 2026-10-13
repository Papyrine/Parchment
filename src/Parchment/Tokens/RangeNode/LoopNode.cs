sealed record LoopNode(
    string OpenAnchorName,
    string CloseAnchorName,
    string LoopVariable,
    Expression LoopSource,
    // First dotted identifier of LoopSource (e.g. `Customer.Lines`), or null when the source is not
    // a plain member path. Precomputed at tree-build time so the render path can probe the editable
    // collection map without re-walking the AST via IdentifierVisitor.Collect on every loop.
    string? SourceReference,
    // The `limit:` / `offset:` / `reversed` modifiers of the tag, as Fluid parsed them. Limit and
    // Offset are expressions rather than numbers because liquid allows `limit: Max`.
    Expression? Limit,
    Expression? Offset,
    bool Reversed,
    // True when a paragraph between the tags anchors a text box that holds a token. That paragraph
    // carries anchors below its direct children, which is as far as the runner looks unless told
    // otherwise (ScopeTreeRunner.CollectAnchors). Precomputed at tree-build time, so a block
    // without one, which is nearly every block, costs the render a bool check.
    bool NestedAnchors,
    IReadOnlyList<RangeNode> Body) :
    RangeNode;
