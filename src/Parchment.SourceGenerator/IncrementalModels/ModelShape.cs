/// <summary>
/// Flat snapshot of a reachable model type graph. <see cref="Types"/> is ordered by BFS from
/// <see cref="RootTypeFullyQualifiedName"/> so equality is deterministic.
/// </summary>
sealed record ModelShape(
    string RootTypeFullyQualifiedName,
    EquatableArray<TypeEntry> Types,
    // Diagnostic ids the generated file disables, sorted: the ones a reference to a member or type
    // of the graph raises at the use site ([Obsolete], [Experimental], the trimming attributes).
    EquatableArray<string> SuppressedDiagnostics = default);
