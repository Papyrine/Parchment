namespace Parchment;

/// <summary>
/// Marks a model property or field whose value should be rendered as a Word table by
/// <c>Excelsior.WordTableBuilder</c> when referenced via a <c>{{ Property }}</c> substitution
/// token. The member must be an <see cref="System.Collections.Generic.IEnumerable{T}"/>; element
/// columns, headings, ordering, and formatting are then derived from the element type's
/// <c>[Column]</c> attributes per Excelsior's normal conventions.
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class ExcelsiorTableAttribute :
    Attribute
{
    /// <summary>
    /// Optional Word paragraph style id applied to every header cell paragraph. The style must be
    /// defined in the host document's styles part. Lets a branded template drive the header font,
    /// size, colour, and spacing. Maps to <c>WordTableBuilder.HeadingParagraphStyle</c>.
    /// </summary>
    public string? HeadingParagraphStyle { get; set; }

    /// <summary>
    /// Optional Word paragraph style id applied to every data cell paragraph — including
    /// <c>IsHtml</c> and link cells. The style must be defined in the host document's styles part.
    /// Lets a branded template drive the body font, size, and spacing. Maps to
    /// <c>WordTableBuilder.BodyParagraphStyle</c>.
    /// </summary>
    public string? BodyParagraphStyle { get; set; }
}
