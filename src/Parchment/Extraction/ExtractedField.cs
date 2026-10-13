namespace Parchment;

/// <summary>
/// One editable field read from a document. <paramref name="Path"/> is the dotted model path
/// (the content control's tag); for a member of an editable collection's row it carries the
/// collection and the row index — <c>Budgets[1].Amount</c>. <paramref name="Value"/> is typed as the model member
/// (string / bool / Date / enum / numeric) when <paramref name="State"/> is
/// <see cref="FieldState.Extracted"/> or <see cref="FieldState.Unreachable"/>, otherwise null. <paramref name="RawText"/> is the
/// control's visible text, kept for diagnostics — most useful on
/// <see cref="FieldState.ParseFailed"/>.
/// </summary>
public sealed record ExtractedField(
    string Path,
    FieldState State,
    object? Value,
    string? RawText);
