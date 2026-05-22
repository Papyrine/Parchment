namespace Parchment.Generated;

public sealed record ExcelsiorTableMapEntry(
    string DottedPath,
    Type ElementType,
    Func<object, object?> Getter,
    string? HeadingParagraphStyle = null,
    string? BodyParagraphStyle = null);