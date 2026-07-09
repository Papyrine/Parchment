/// <summary>
/// One <c>[EditableField]</c>-marked collection reachable from the root model — rendered as a Word
/// repeating section (`w15:repeatingSection`) so users can add / remove / reorder items in Word.
/// <see cref="ElementMap"/> holds the element type's editable members as <em>item-relative</em>
/// entries (dotted paths, getters and setters rooted at an element instance); those relative paths
/// are the control tags inside each repeated item. <see cref="Setter"/> writes the rebuilt list back
/// onto the model; <see cref="ElementFactory"/> constructs a fresh element during extraction.
/// </summary>
sealed record CollectionEntry(
    string DottedPath,
    Type ElementType,
    Action<object, object?> Setter,
    Func<object, bool> CanReach,
    Func<object> ElementFactory,
    EditableMap ElementMap,
    bool IsArray)
{
    /// <summary>Friendly name surfaced in the repeating-section control chrome (<c>w:alias</c>).</summary>
    public string Alias { get; } = DottedPath[(DottedPath.LastIndexOf('.') + 1)..];
}
