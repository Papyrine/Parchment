namespace Parchment.Generated;

/// <summary>
/// Pre-compiled editable-collection registration data emitted by the source generator — the
/// repeating-section counterpart of <see cref="EditableFieldMapEntry"/>. Public data carrier for
/// <c>GeneratedRegistration.RegisterEditableCollections</c>; not intended for hand-written use.
/// <paramref name="ElementFields"/> are the element type's editable members with getters/setters
/// rooted at an <em>element</em> instance (dotted paths item-relative). <paramref name="Setter"/>
/// writes the rebuilt list onto the root model; <paramref name="ElementFactory"/> constructs a fresh
/// element during extraction; <paramref name="IsArray"/> selects <c>TItem[]</c> over
/// <c>List&lt;TItem&gt;</c> when rebuilding.
/// </summary>
public sealed record CollectionFieldMapEntry(
    string DottedPath,
    Type ElementType,
    Action<object, object?> Setter,
    Func<object, bool> CanReach,
    Func<object> ElementFactory,
    IReadOnlyList<EditableFieldMapEntry> ElementFields,
    bool IsArray);
