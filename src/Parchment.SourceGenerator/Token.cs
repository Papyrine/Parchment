public sealed record Token(
    TokenKind Kind,
    string Source,
    IReadOnlyList<IReadOnlyList<string>> References,
    string? LoopVariable,
    bool HasOtherContent,
    bool IsPlainIdentifier = false,
    // ForOpen only. References holds the paths of the loop source alone; the member paths inside
    // `limit:` / `offset:` are kept apart so the source stays References[0] whatever the tag carries.
    IReadOnlyList<IReadOnlyList<string>>? LoopModifierReferences = null,
    bool HasLoopModifiers = false,
    int Container = 0,
    // A `for` over a range such as `(1..Count)`. References are then the range's bounds rather
    // than the thing iterated, and the loop variable is untyped.
    bool LoopSourceIsRange = false,
    // The token starts inside a content control that holds a single typed value (plain text, date
    // picker, drop-down list, combo box, checkbox, picture). Set only when the scan is given spans.
    bool InValueControl = false);
