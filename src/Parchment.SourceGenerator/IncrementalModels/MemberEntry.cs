sealed record MemberEntry(
    string Name,
    string TypeFullyQualifiedName,
    bool IsExcelsiorTable = false,
    bool IsHtml = false,
    bool IsMarkdown = false,
    bool IsStringList = false,
    bool IsStatic = false,
    string? ExcelsiorHeadingParagraphStyle = null,
    string? ExcelsiorBodyParagraphStyle = null);
