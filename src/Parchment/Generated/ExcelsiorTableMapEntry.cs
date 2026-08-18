namespace Parchment.Generated;

public sealed record ExcelsiorTableMapEntry(
    string DottedPath,
    Type ElementType,
    Func<object, object?> Getter,
    string? HeadingParagraphStyle = null,
    string? BodyParagraphStyle = null,
    string? TableStyle = null,
    // The [ExcelsiorTable(Configure = ...)] method, closed over the element type by the generated
    // code - the builder it receives is a WordTableBuilder<TElement> behind the object.
    Action<object>? Configure = null);