namespace Parchment.Generated;

/// <summary>
/// The Word content-control flavour an <c>[EditableField]</c> member renders as. Public because
/// source-generator-emitted registration code references it; not intended for hand-written use.
/// </summary>
public enum EditableFieldKind
{
    /// <summary>Plain-text control (<c>w:text</c>) for string members.</summary>
    Text,

    /// <summary>Plain-text control whose text is parsed numerically at extraction.</summary>
    Number,

    /// <summary>Checkbox control (<c>w14:checkbox</c>) for bool members.</summary>
    Checkbox,

    /// <summary>
    /// Date-picker control (<c>w:date</c>) carrying a canonical <c>w:fullDate</c>. Used for
    /// <c>DateOnly</c> and <c>DateTime</c> members.
    /// </summary>
    Date,

    /// <summary>
    /// Plain-text control (<c>w:text</c>) holding a round-trippable ISO 8601 value for
    /// <c>DateTimeOffset</c> members. Word has no offset-aware picker, and <c>w:fullDate</c>
    /// cannot carry an offset, so the run text is the source of truth.
    /// </summary>
    DateTimeOffset,

    /// <summary>
    /// Plain-text control (<c>w:text</c>) holding a time value for <c>TimeOnly</c> members.
    /// Word has no time-only picker, so the run text is the source of truth.
    /// </summary>
    Time,

    /// <summary>Dropdown control (<c>w:dropDownList</c>) with one item per enum member.</summary>
    DropDown
}
