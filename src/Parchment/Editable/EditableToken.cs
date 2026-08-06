/// <summary>
/// Marker returned by <c>ScopeTreeRunner.TryResolveEditable</c> so the substitution loop can
/// route the token through the editable-field splice instead of plain text replacement.
/// Internal-only — never surfaced to user code (unlike <see cref="Parchment.TokenValue"/>).
/// </summary>
sealed record EditableToken(EditableEntry Entry, object? Value);
