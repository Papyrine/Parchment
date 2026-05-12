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
    Attribute;
