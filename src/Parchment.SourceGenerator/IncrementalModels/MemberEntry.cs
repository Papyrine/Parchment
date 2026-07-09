sealed record MemberEntry(
    string Name,
    string TypeFullyQualifiedName,
    bool IsExcelsiorTable = false,
    bool IsHtml = false,
    bool IsMarkdown = false,
    bool IsStringList = false,
    bool IsStatic = false,
    string? ExcelsiorHeadingParagraphStyle = null,
    string? ExcelsiorBodyParagraphStyle = null,
    bool IsEditable = false,
    EditableFieldKind? EditableKind = null,
    bool EditableIsNullable = false,
    bool HasUsableSetter = false,
    bool EditableMultiLine = false,
    string? EditableDateFormat = null,
    // Non-null when this is an [EditableField] collection of a POCO element type (rendered as a Word
    // repeating section). EditableKind stays null; the value is the element type's fully-qualified name.
    string? EditableCollectionElementFqn = null);
