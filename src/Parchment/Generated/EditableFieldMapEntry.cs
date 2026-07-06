namespace Parchment.Generated;

/// <summary>
/// Pre-compiled editable-field registration data emitted by the source generator. Public data
/// carrier for <c>GeneratedRegistration.RegisterEditable</c>; not intended for hand-written use.
/// <paramref name="Getter"/> reads the value from the root model at render time;
/// <paramref name="Setter"/> writes an extracted value back; <paramref name="CanReach"/> reports
/// whether every intermediate object on the path is non-null (so extraction can validate before
/// mutating).
/// </summary>
public sealed record EditableFieldMapEntry(
    string DottedPath,
    EditableFieldKind Kind,
    Type ClrType,
    bool IsNullable,
    Func<object, object?> Getter,
    Action<object, object?> Setter,
    Func<object, bool> CanReach,
    bool MultiLine,
    string? DateFormat);
