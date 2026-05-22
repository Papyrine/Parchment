sealed record ExcelsiorTableEntry(
    Type ElementType,
    Func<object, object?> Getter,
    string? HeadingParagraphStyle = null,
    string? BodyParagraphStyle = null);