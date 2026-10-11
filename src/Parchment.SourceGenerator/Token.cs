public sealed record Token(
    TokenKind Kind,
    string Source,
    IReadOnlyList<IReadOnlyList<string>> References,
    string? LoopVariable,
    bool HasOtherContent,
    bool IsPlainIdentifier = false,
    int Container = 0);
