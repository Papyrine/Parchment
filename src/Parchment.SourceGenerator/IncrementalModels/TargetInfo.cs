/// <summary>
/// Primitive-only snapshot of a <c>[ParchmentModel]</c>-decorated class. Every field is
/// value-equatable so this can flow through the incremental pipeline without pinning to a
/// specific <see cref="Compilation"/>. In particular, no <see cref="ISymbol"/> fields —
/// symbols change identity between compilations and would invalidate the cache on every edit.
///
/// The decorated class IS the binding model — <see cref="ModelFullyQualifiedName"/> and
/// <see cref="ModelDisplayName"/> describe the same symbol as <see cref="DeclaringName"/>.
/// <see cref="ModelDisplayName"/> includes the enclosing-type chain (e.g.
/// <c>"TravelRequestBriefGenerator.Info"</c>) so PARCH001 messages disambiguate models that
/// share a simple name across different enclosing types.
///
/// <see cref="EnclosingTypes"/> is outermost-to-innermost; an empty array means the target sits
/// at namespace scope. <see cref="ExtractError"/> is non-null when the extract step found a
/// blocker (e.g. a non-partial enclosing type) — Process turns it into a PARCH011 diagnostic
/// and skips registration generation.
/// </summary>
sealed record TargetInfo(
    string? DeclaringNamespace,
    string DeclaringName,
    string DeclaringKind,
    EquatableArray<EnclosingType> EnclosingTypes,
    string ModelFullyQualifiedName,
    string ModelDisplayName,
    string TemplatePath,
    ProtectionMode Protection,
    EquatableLocation Location,
    ModelShape Shape,
    string? ExtractError);
