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

    /// <summary>Date-picker control (<c>w:date</c>) carrying a canonical <c>w:fullDate</c>.</summary>
    Date,

    /// <summary>Dropdown control (<c>w:dropDownList</c>) with one item per enum member.</summary>
    DropDown
}
