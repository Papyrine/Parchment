namespace Parchment;

/// <summary>
/// Marks a model property or field as user-editable in the rendered docx. The substitution site
/// is wrapped in a Word content control (<c>w:sdt</c>) tagged with the member's dotted path, and
/// an editable-range exception (<c>w:permStart</c>/<c>w:permEnd</c>, everyone) is punched around
/// it. When the registered template contains at least one editable field, the output document is
/// locked read-only via <c>w:documentProtection</c> — see <see cref="ProtectionMode"/>.
/// Values entered by users are read back via <c>ParchmentExtractor</c>.
///
/// Supported member types: <c>string</c>, <c>bool</c> (checkbox), <c>DateOnly</c> /
/// <c>DateTime</c> (date picker), <c>DateTimeOffset</c> / <c>TimeOnly</c> (round-trippable plain
/// text — Word has no offset or time-only picker), enums (dropdown), and the numeric primitives /
/// <c>decimal</c>. Nullable variants are supported except <c>bool?</c> — a checkbox cannot
/// represent null. The member must have a public non-init setter so extraction can write back.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class EditableFieldAttribute :
    Attribute
{
    /// <summary>
    /// For string members: allow multi-line input (<c>w:text w:multiLine="1"</c>). Newlines in
    /// the rendered value become soft line breaks; extraction maps breaks back to <c>\n</c>.
    /// Non-multiline string values have newlines collapsed to spaces.
    /// </summary>
    public bool MultiLine { get; set; }

    /// <summary>
    /// For date members: the display format of the date picker (<c>w:dateFormat</c>), also used
    /// to format the rendered text. Defaults to <c>yyyy-MM-dd</c>. The canonical value read back
    /// by extraction comes from the control's <c>w:fullDate</c>, so the display format does not
    /// need to be round-trippable.
    /// </summary>
    public string? DateFormat { get; set; }
}
