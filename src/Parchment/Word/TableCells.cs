/// <summary>
/// Closes every table cell with a paragraph, which Word requires and the schema does not.
/// </summary>
/// <remarks>
/// <para>
/// <c>CT_Tc</c> allows a cell with no block content at all, so <c>OpenXmlValidator</c> passes an
/// empty <c>w:tc</c> and one whose last block is a nested <c>w:tbl</c>. Word accepts neither: it
/// reports the file as corrupt and refuses to open it.
/// </para>
/// <para>
/// A render reaches both shapes by removing or replacing the only paragraph a cell had — a false
/// <c>{% if %}</c>, a <c>{% for %}</c> over nothing, a structural token that produced no elements,
/// or one that produced a table. Those are many sites, each of which would have to know whether its
/// host was the last paragraph of a cell, so the repair is one pass over the part once they have
/// all run.
/// </para>
/// </remarks>
static class TableCells
{
    /// <summary>
    /// Appends an empty paragraph to every cell under <paramref name="root"/> that lacks a final one.
    /// </summary>
    public static void EnsureFinalParagraph(OpenXmlCompositeElement root)
    {
        // Collected before appending: mutating the tree while enumerating Descendants is not safe.
        // A part with no cell to repair allocates nothing.
        List<TableCell>? open = null;
        foreach (var cell in root.Descendants<TableCell>())
        {
            if (!EndsWithParagraph(cell))
            {
                open ??= [];
                open.Add(cell);
            }
        }

        if (open == null)
        {
            return;
        }

        foreach (var cell in open)
        {
            cell.AppendChild(new Paragraph());
        }
    }

    static bool EndsWithParagraph(TableCell cell)
    {
        for (var child = cell.LastChild; child != null; child = child.PreviousSibling())
        {
            // A block-level content control or custom xml wrapper carries the paragraph inside it,
            // which is how Word itself writes a cell whose content is one control.
            if (child is Paragraph or SdtBlock or CustomXmlBlock)
            {
                return true;
            }

            if (child is Table or TableCellProperties)
            {
                return false;
            }

            // Anything else is a range marker — a bookmark or permission end — which Word allows
            // after the closing paragraph.
        }

        return false;
    }
}
