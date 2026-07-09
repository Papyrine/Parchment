/// <summary>
/// SG-side copy of <c>Parchment.Generated.EditableFieldKind</c> — the generator cannot reference
/// Parchment.dll. Values must stay in lockstep; emission writes the runtime enum by member name.
/// </summary>
enum EditableFieldKind
{
    Text,
    Number,
    Checkbox,
    Date,
    DateTimeOffset,
    Time,
    DropDown,
    Html
}
