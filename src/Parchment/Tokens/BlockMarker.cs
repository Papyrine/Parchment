sealed record BlockMarker(
    BlockTagKind Kind,
    string Source,
    Expression? Condition,
    string? LoopVariable,
    Expression? LoopSource,
    Expression? LoopLimit = null,
    Expression? LoopOffset = null,
    bool LoopReversed = false);
