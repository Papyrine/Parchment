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
/// <para>
/// A block-level content control or custom xml element is not a block of its own. Word reads
/// through it, so a cell whose control ends with a table ends with a table, and is refused like
/// any other. The paragraph that closes such a cell is appended to the cell, after the control: put
/// inside, it would become part of the control's content, which for an editable field is the value
/// read back at extraction.
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
        var last = LastBlock(cell);
        if (last == null)
        {
            // No paragraph and no table at any depth. Word refuses a cell with nothing in it, but
            // opens one whose only block is a content control with no content, which is what a
            // false {% if %} inside a template's own control leaves. That one stays as it is.
            return cell.GetFirstChild<SdtBlock>() != null;
        }

        return last is Paragraph;
    }

    /// <summary>
    /// The last paragraph or table under <paramref name="container"/>, read through the block-level
    /// wrappers the way Word reads them, or null when it holds neither.
    /// </summary>
    static OpenXmlElement? LastBlock(OpenXmlElement container)
    {
        for (var child = container.LastChild; child != null; child = child.PreviousSibling())
        {
            if (child is Paragraph or Table)
            {
                return child;
            }

            // A content control or custom xml wrapper ends with whatever its content ends with.
            // One that carries the paragraph inside is how Word itself writes a cell whose content
            // is one control; one left ending with a table has not closed the cell.
            var content = WrappedContent(child);
            if (content == null)
            {
                // Anything else is a range marker — a bookmark or permission end — which Word
                // allows after the closing paragraph, or the properties of the cell or wrapper.
                continue;
            }

            var inner = LastBlock(content);
            if (inner != null)
            {
                return inner;
            }
        }

        return null;
    }

    static OpenXmlElement? WrappedContent(OpenXmlElement element)
    {
        if (element is SdtBlock control)
        {
            return control.SdtContentBlock;
        }

        if (element is CustomXmlBlock)
        {
            return element;
        }

        return null;
    }
}
