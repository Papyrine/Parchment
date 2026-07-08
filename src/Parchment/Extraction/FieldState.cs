namespace Parchment;

/// <summary>
/// Outcome of reading one editable field from a document. Only <see cref="Extracted"/> values
/// (and <see cref="Empty"/> for nullable members) are written back by
/// <c>ExtractResult&lt;TModel&gt;.ApplyTo</c>; the other states surface problems for the caller
/// to inspect.
/// </summary>
public enum FieldState
{
    /// <summary>The control held a value that parsed as the member's type.</summary>
    Extracted,

    /// <summary>
    /// The control shows its placeholder (or was cleared). Applied as null to nullable members;
    /// skipped for non-nullable members.
    /// </summary>
    Empty,

    /// <summary>
    /// No content control with the member's tag exists in the document — the control was
    /// deleted, or the document was not produced from a template bound to this model.
    /// </summary>
    Missing,

    /// <summary>
    /// The control's content could not be parsed as the member's type (see
    /// <c>ExtractedField.RawText</c>). Typically a culture mismatch or free-typed text in a
    /// numeric / date field.
    /// </summary>
    ParseFailed,

    /// <summary>
    /// A second control with an already-seen tag. The first occurrence won; this one is ignored.
    /// </summary>
    Duplicate
}
