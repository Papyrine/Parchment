namespace Parchment;

/// <summary>
/// Controls whether a registered docx template's output is locked down with
/// <c>w:documentProtection w:edit="readOnly"</c> (with editable-range exceptions around each
/// <see cref="EditableFieldAttribute"/> site). Protection is cooperative, not security: Word
/// enforces it in the UI, but anyone can remove it (no password is set) or edit the underlying
/// XML directly.
/// </summary>
public enum ProtectionMode
{
    /// <summary>
    /// Apply read-only protection when the model declares at least one
    /// <see cref="EditableFieldAttribute"/> member. Templates without editable fields are left
    /// unprotected. This is the default.
    /// </summary>
    WhenEditable,

    /// <summary>
    /// Never apply protection. Editable fields still render as tagged content controls (and
    /// still extract), but the rest of the document remains editable too.
    /// </summary>
    None
}
