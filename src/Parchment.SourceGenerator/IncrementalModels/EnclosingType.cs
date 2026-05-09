/// <summary>
/// One link in the enclosing-type chain of a <c>[ParchmentTemplate]</c>-decorated nested class.
/// Stored outermost-to-innermost on <see cref="TargetInfo.EnclosingTypes"/> so the generator
/// can wrap the emitted partial in matching <c>partial {Kind} {Name} { ... }</c> declarations.
/// Both fields are primitive strings so <see cref="EquatableArray{T}"/> equality stays cheap and
/// the incremental pipeline doesn't churn on enclosing-type symbol identity changes.
/// </summary>
sealed record EnclosingType(string Name, string Kind);
