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

    /// <summary>
    /// Optional Word table style id applied to the table, in place of the built-in
    /// <c>TableGrid</c>. The style must be defined in the host document's styles part — a
    /// template's own style is the template's to define. Lets a branded template drive the table's
    /// borders, banding, and cell margins. Maps to <c>WordTableBuilder.TableStyle</c>.
    /// </summary>
    public string? TableStyle { get; set; }

    /// <summary>
    /// Optional name of a static method on the declaring type - written as <c>nameof(...)</c> -
    /// that receives the <c>WordTableBuilder&lt;TElement&gt;</c> before the table is built:
    /// <c>static void Configure(WordTableBuilder&lt;Line&gt; builder)</c>. The escape hatch for
    /// everything the attribute cannot say - per-column configuration, merged rows - since an
    /// attribute can only carry constants and the builder is otherwise Parchment's to construct.
    /// </summary>
    /// <remarks>
    /// Resolved at compile time: the source generator emits a direct call, so a missing method is
    /// PARCH026 rather than a render-time failure, and a wrong signature is an ordinary compile
    /// error. Runs after the attribute's own settings, so what it sets wins.
    /// </remarks>
    public string? Configure { get; set; }
}
