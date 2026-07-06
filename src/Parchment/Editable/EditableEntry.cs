/// <summary>
/// One editable member reachable from the root model. <see cref="Getter"/> reads the current
/// value at render time; <see cref="Setter"/> writes an extracted value back onto a model
/// instance; <see cref="CanReach"/> reports whether every intermediate object on the dotted path
/// is non-null so extraction can validate before mutating.
/// </summary>
sealed record EditableEntry(
    string DottedPath,
    EditableFieldKind Kind,
    Type ClrType,
    bool IsNullable,
    Func<object, object?> Getter,
    Action<object, object?> Setter,
    Func<object, bool> CanReach,
    bool MultiLine,
    string? DateFormat)
{
    /// <summary>
    /// Friendly name surfaced in Word's content-control chrome (<c>w:alias</c>) — the leaf
    /// member name.
    /// </summary>
    public string Alias { get; } = DottedPath[(DottedPath.LastIndexOf('.') + 1)..];
}

/// <summary>
/// Marker returned by <c>ScopeTreeRunner.TryResolveEditable</c> so the substitution loop can
/// route the token through the editable-field splice instead of plain text replacement.
/// Internal-only — never surfaced to user code (unlike <see cref="Parchment.TokenValue"/>).
/// </summary>
sealed record EditableToken(EditableEntry Entry, object? Value);
