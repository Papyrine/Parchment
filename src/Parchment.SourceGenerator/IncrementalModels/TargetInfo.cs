/// <summary>
/// Primitive-only snapshot of a <c>[ParchmentTemplate]</c>-decorated class. Every field is
/// value-equatable so this can flow through the incremental pipeline without pinning to a
/// specific <see cref="Compilation"/>. In particular, no <see cref="ISymbol"/> fields —
/// symbols change identity between compilations and would invalidate the cache on every edit.
///
/// <see cref="EnclosingTypes"/> is outermost-to-innermost; an empty array means the target sits
/// at namespace scope. <see cref="ExtractError"/> is non-null when the extract step found a
/// blocker (e.g. a non-partial enclosing type) — Process turns it into a PARCH011 diagnostic
/// and skips registration generation.
/// </summary>
sealed record TargetInfo(
    string? DeclaringNamespace,
    string DeclaringName,
    EquatableArray<EnclosingType> EnclosingTypes,
    string ModelFullyQualifiedName,
    string ModelSimpleName,
    string TemplatePath,
    EquatableLocation Location,
    ModelShape Shape,
    string? ExtractError);
