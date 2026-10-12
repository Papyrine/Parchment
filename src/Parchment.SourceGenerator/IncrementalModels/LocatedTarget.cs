/// <summary>
/// A <see cref="TargetInfo"/> together with the syntax tree its attribute was read from.
///
/// The tree is what lets a diagnostic be reported at a source location, which is the only kind
/// <c>#pragma warning disable</c>, <c>[SuppressMessage]</c> and a per-file <c>.editorconfig</c>
/// severity apply to. It compares by reference. That is safe for caching where a
/// <see cref="Compilation"/> or an <see cref="ISymbol"/> would not be: a compilation hands back
/// the same tree instance for every file an edit did not touch, so this value only changes when
/// the file declaring the model does.
/// </summary>
sealed record LocatedTarget(TargetInfo Target, SyntaxTree? Tree);
