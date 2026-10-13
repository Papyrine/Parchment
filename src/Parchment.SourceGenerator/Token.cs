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
    // UnknownBlock only: why a tag that is otherwise supported was refused, such as an `else`
    // carrying arguments. Null for a tag that is not supported at all.
    string? Problem = null,
    // Substitution only, and set together: a filter that produces a token, and the filter applied
    // after it. See TokenFilterOrder.
    string? TokenFilter = null,
    string? FilterAfterTokenFilter = null,
    // The token starts inside a content control that holds a single typed value (plain text, date
    // picker, drop-down list, combo box, checkbox, picture). Set only when the scan is given spans.
    bool InValueControl = false,
    // The token's paragraph is in an mc:Fallback, the VML copy of a text box. Only set by the scan
    // of the body the editable rules run over.
    bool InFallback = false,
    // The index of the part the token's paragraph sits in. The runtime builds one scope tree per
    // part, so block tags only pair within a part.
    int Part = 0,
    // Substitution only: the error Fluid gave for a token it could not parse.
    string? ParseError = null);
