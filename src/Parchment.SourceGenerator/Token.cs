public sealed record Token(
    TokenKind Kind,
    string Source,
    IReadOnlyList<IReadOnlyList<string>> References,
    string? LoopVariable,
    bool HasOtherContent,
    bool IsPlainIdentifier = false,
    // A `for` over a range such as `(1..Count)`. References are then the range's bounds rather
    // than the thing iterated, and the loop variable is untyped.
    bool LoopSourceIsRange = false);
