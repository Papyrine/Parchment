sealed record MemberEntry(
    string Name,
    string TypeFullyQualifiedName,
    bool IsExcelsiorTable = false,
    bool IsHtml = false,
    bool IsMarkdown = false,
    // [Html] + [Markdown] on one member, or either contradicted by [StringSyntax] — reported as
    // PARCH023 and excluded from format-map emission.
    bool HasFormatConflict = false,
    // [Html]/[Markdown]/[StringSyntax] on a member already typed as a TokenValue — reported as
    // PARCH025 and excluded from format-map emission, since the type declares the same thing.
    bool FormatMarkerOnTokenValue = false,
    bool IsStringList = false,
    bool IsStatic = false,
    string? ExcelsiorHeadingParagraphStyle = null,
    string? ExcelsiorBodyParagraphStyle = null,
    string? ExcelsiorTableStyle = null,
    bool IsEditable = false,
    EditableFieldKind? EditableKind = null,
    bool EditableIsNullable = false,
    bool HasUsableSetter = false,
    bool EditableMultiLine = false,
    string? EditableDateFormat = null,
    // Non-null when this is an [EditableField] collection of a POCO element type (rendered as a Word
    // repeating section). EditableKind stays null; the value is the element type's fully-qualified name.
    string? EditableCollectionElementFqn = null,
    // The fully-qualified [ExcelsiorTable(Configure = ...)] call target - declaring type dot method -
    // or null when the attribute names none. When the name resolves to no static method the flag is
    // set and this holds the raw name for the diagnostic; emission skips it, so the clean report is
    // the only error rather than the first of a cascade.
    string? ExcelsiorConfigure = null,
    bool ExcelsiorConfigureMissing = false);
