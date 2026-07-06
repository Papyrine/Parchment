/// <summary>
/// Per-render allocator for the ids shared by editable-field elements: <c>w:sdt</c> ids and the
/// <c>w:permStart</c>/<c>w:permEnd</c> pair ids. Sequential in processing order, which keeps
/// renders deterministic (same template + same model → same ids) without any randomness.
/// </summary>
sealed class EditableState
{
    int next = 1;

    public int NextId() => next++;
}
