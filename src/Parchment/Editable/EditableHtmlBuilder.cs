using SdtLock = DocumentFormat.OpenXml.Wordprocessing.Lock;

/// <summary>
/// Builds the block-level editable-HTML region. The block content control (a rich-text
/// <c>w:sdt</c> — no kind element, which is what Word treats as rich text) carries the dotted
/// model path as its <c>w:tag</c> (the round-trip key) and <c>w:lock="sdtLocked"</c> so users can
/// edit the formatted content but not delete the control. A perm range (edGrp="everyone") punches
/// an editable exception through the document's read-only protection — the block-level analogue
/// of <see cref="EditableFieldBuilder"/>. Rendered HTML blocks (from OpenXmlHtml's
/// <c>WordHtmlConverter</c>) are moved into the control's <c>sdtContent</c>; an empty value
/// renders the grey placeholder paragraph.
///
/// The perm range takes one of two shapes:
/// <list type="bullet">
/// <item><c>[w:permStart, w:sdt(block), w:permEnd]</c> — the general body shape.</item>
/// <item>When the control is the sole content of a table cell, a whole-cell range instead:
/// row-level <c>w:permStart w:colFirst/w:colLast</c> + <c>w:permEnd</c> (the markup Word's own UI
/// emits for cell exceptions). Paragraph formatting (lists, indents) touches the paragraph mark,
/// and for the cell's last paragraph the end-of-cell marker — with a tight block-level range those
/// sit on/outside the boundary, so Word silently refuses list operations on the first/last
/// paragraphs and strands AutoFormat indents. A whole-cell range contains the end-of-cell marker,
/// making paragraph formatting work on every paragraph. Applied only when nothing else shares the
/// cell, so the wider range exposes no sibling content to editing.</item>
/// </list>
/// </summary>
static class EditableHtmlBuilder
{
    const string placeholderText = "Enter rich text";

    public static IReadOnlyList<OpenXmlElement> Build(
        EditableEntry entry,
        IReadOnlyList<OpenXmlElement> content,
        EditableState state,
        Paragraph host)
    {
        var id = state.NextId();

        var sdtPr = new SdtProperties();
        sdtPr.AppendChild(new SdtAlias
        {
            Val = entry.Alias
        });
        sdtPr.AppendChild(new Tag
        {
            Val = entry.DottedPath
        });
        sdtPr.AppendChild(new SdtId
        {
            Val = id
        });
        sdtPr.AppendChild(new SdtLock
        {
            Val = LockingValues.SdtLocked
        });

        var isPlaceholder = content.Count == 0;
        if (isPlaceholder)
        {
            sdtPr.AppendChild(new ShowingPlaceholder());
        }

        var contentBlock = new SdtContentBlock();
        if (isPlaceholder)
        {
            contentBlock.AppendChild(PlaceholderParagraph());
        }
        else
        {
            foreach (var element in content)
            {
                contentBlock.AppendChild(element);
            }
        }

        var sdt = new SdtBlock(sdtPr, contentBlock);

        if (TryGetSoleCell(host, out var cell, out var row))
        {
            InstallWholeCellPermRange(row, cell, id);
            return [sdt];
        }

        return
        [
            new PermStart
            {
                Id = id,
                EditorGroup = RangePermissionEditingGroupValues.Everyone
            },
            sdt,
            new PermEnd
            {
                Id = id
            }
        ];
    }

    /// <summary>
    /// True when <paramref name="host"/> is the only content of a table cell (everything else is
    /// cell properties), so replacing it with the control leaves the control as the cell's entire
    /// content and a whole-cell perm range exposes nothing extra.
    /// </summary>
    static bool TryGetSoleCell(
        Paragraph host,
        [NotNullWhen(true)] out TableCell? cell,
        [NotNullWhen(true)] out TableRow? row)
    {
        cell = null;
        row = null;
        if (host.Parent is not TableCell parentCell ||
            parentCell.Parent is not TableRow parentRow)
        {
            return false;
        }

        foreach (var child in parentCell.ChildElements)
        {
            if (child != host &&
                child is not TableCellProperties)
            {
                return false;
            }
        }

        cell = parentCell;
        row = parentRow;
        return true;
    }

    /// <summary>
    /// Marks the whole cell editable: <c>w:permStart</c> (with <c>w:colFirst</c>/<c>w:colLast</c>
    /// selecting the cell's grid columns) before the row's cells and the matching
    /// <c>w:permEnd</c> after them. The row itself is never structurally replaced, so mutating it
    /// here — while the host paragraph is still queued for replacement — is safe.
    /// </summary>
    static void InstallWholeCellPermRange(TableRow row, TableCell cell, int id)
    {
        var colFirst = 0;
        foreach (var sibling in row.Elements<TableCell>())
        {
            if (sibling == cell)
            {
                break;
            }

            colFirst += GridSpan(sibling);
        }

        var start = new PermStart
        {
            Id = id,
            EditorGroup = RangePermissionEditingGroupValues.Everyone,
            ColumnFirst = colFirst,
            ColumnLast = colFirst + GridSpan(cell) - 1
        };
        row.InsertBefore(start, row.Elements<TableCell>().First());
        row.AppendChild(new PermEnd
        {
            Id = id
        });
    }

    static int GridSpan(TableCell cell) =>
        cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;

    static Paragraph PlaceholderParagraph()
    {
        // Word's built-in grey placeholder look. The style may not exist in the template's styles
        // part — Word tolerates the dangling reference and falls back to plain text.
        var runProperties = new RunProperties(
            new RunStyle
            {
                Val = "PlaceholderText"
            });
        var run = new Run(
            runProperties,
            new Text(placeholderText)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        return new(run);
    }
}
