/// <summary>
/// One link in the enclosing-type chain of a <c>[ParchmentModel]</c>-decorated nested class.
/// Stored outermost-to-innermost on <see cref="TargetInfo.EnclosingTypes"/> so the generator
/// can wrap the emitted partial in matching <c>partial {Kind} {Name} { ... }</c> declarations.
/// Every field is a primitive so <see cref="EquatableArray{T}"/> equality stays cheap and
/// the incremental pipeline doesn't churn on enclosing-type symbol identity changes.
/// <see cref="IsModuleAccessible"/> is false for a private or protected declaration, which a
/// module initializer cannot sit inside.
/// </summary>
sealed record EnclosingType(string Name, string Kind, bool IsModuleAccessible);
