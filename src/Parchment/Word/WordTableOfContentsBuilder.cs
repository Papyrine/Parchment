/// <summary>
/// Fills a table-of-contents field with the entries it describes, so the document reaches its final
/// length before anything measures it.
/// </summary>
/// <remarks>
/// Word builds these entries itself when it updates the field. Building them here instead is what
/// makes a resolved table of contents possible at all: entries added after pagination would push
/// every heading below them onto a later page and invalidate the numbers just measured, whereas a
/// number written into a line that already exists reflows nothing.
/// <para>
/// The entries stay inside the field rather than beside it. A reader who refreshes the field then
/// gets them replaced, which is the point — outside it, refreshing would insert a second table.
/// </para>
/// </remarks>
static partial class WordTableOfContentsBuilder
{
    // Stands in for the page number until the resolver reports one. A single digit keeps the entry
    // the width it will end up, so filling the real number in cannot rewrap the line.
    internal const string PagePlaceholder = "1";

    public static void WriteEntries(Body body, WordFieldRegion field, BookmarkIds ids)
    {
        var maxLevel = ParseMaxLevel(field.Instruction);
        var headings = CollectHeadings(body, maxLevel);
        if (headings.Count == 0)
        {
            // Nothing to list. The placeholder stays, and so does the request for Word to try again.
            return;
        }

        var bookmarks = new WordBookmarkAllocator(ids);
        var tabStop = TextWidth(SectionOf(body, field.Paragraph));
        var entries = new List<Paragraph>();

        // The section's width is the room a field in the body has. One in a table cell has the cell.
        var width = WidthInCells(field.Paragraph, tabStop);
        foreach (var (paragraph, level, text) in headings)
        {
            entries.Add(BuildEntry(bookmarks.Ensure(paragraph), level, text, width));
        }

        // The field's markers move onto the entries: begin/instruction/separate lead the first, and
        // end closes the last, which is how Word writes a table of contents spanning paragraphs.
        var source = field.Paragraph;
        var leading = source.Elements<Run>()
            .TakeWhile(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Separate)
            .Concat(source.Elements<Run>()
                .Where(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate)
                .Take(1))
            .ToList();
        var end = source.Elements<Run>()
            .LastOrDefault(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End);

        // After the entry's properties rather than at the front of it: w:pPr has to stay the
        // paragraph's first child, and Word discards the whole of it when a run comes first — so the
        // first entry alone would lose its TOC style and the tab stop its dot leader hangs off.
        var anchor = entries[0].ParagraphProperties!;
        for (var i = leading.Count - 1; i >= 0; i--)
        {
            leading[i].Remove();
            anchor.InsertAfterSelf(leading[i]);
        }

        if (end != null)
        {
            end.Remove();
            entries[^1].Append(end);
        }

        foreach (var entry in entries)
        {
            source.InsertBeforeSelf(entry);
        }

        source.Remove();
    }

    static Paragraph BuildEntry(string bookmark, int level, string text, int tabStop)
    {
        var properties = new ParagraphProperties
        {
            ParagraphStyleId = new()
            {
                Val = $"TOC{level}"
            }
        };
        properties.Append(
            new Tabs(
                new TabStop
                {
                    Val = TabStopValues.Right,
                    Leader = TabStopLeaderCharValues.Dot,
                    Position = tabStop
                }));

        var link = new Hyperlink
        {
            Anchor = bookmark
        };
        link.Append(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
        link.Append(new Run(new TabChar()));
        link.Append(new Run(new Text(PagePlaceholder)));

        return new(properties, link);
    }

    static List<(Paragraph Paragraph, int Level, string Text)> CollectHeadings(Body body, int maxLevel)
    {
        var found = new List<(Paragraph, int, string)>();

        // Every paragraph rather than the body's own: a heading in a table cell is one level down or
        // more, and Word lists it when it builds the field itself.
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (style == null ||
                !style.StartsWith("Heading", StringComparison.Ordinal) ||
                !int.TryParse(style.AsSpan("Heading".Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out var level) ||
                level > maxLevel)
            {
                continue;
            }

            var text = string.Concat(paragraph.Descendants<Text>().Select(_ => _.Text)).Trim();
            if (text.Length == 0)
            {
                continue;
            }

            found.Add((paragraph, level, text));
        }

        return found;
    }

    static int ParseMaxLevel(string instruction)
    {
        // \o "1-3"
        var match = ParseMaxLevelRegex()
            .Match(instruction);
        if (match.Success &&
            int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var level))
        {
            return level;
        }

        return 3;
    }

    // The section the field sits in, which is not the document's last once [SECTION] has split it: a
    // w:sectPr describes the section that ends at its paragraph, so the one in force here is the
    // first to close at or below the field, and the document-final one only when none does. Looked
    // for among the body's own children, where a section break has to be, starting from whichever
    // of them holds the field — the paragraph itself, or the table it is in a cell of. Read after
    // SectionBreaks.Apply, which is what turns each marker into the full section measured here.
    static SectionProperties? SectionOf(Body body, Paragraph paragraph)
    {
        OpenXmlElement block = paragraph;
        while (block.Parent != null &&
               !ReferenceEquals(block.Parent, body))
        {
            block = block.Parent;
        }

        for (var current = block; current != null; current = current.NextSibling())
        {
            if (current is Paragraph { ParagraphProperties.SectionProperties: { } section })
            {
                return section;
            }
        }

        return body.Elements<SectionProperties>().LastOrDefault();
    }

    // Where the page number sits: hard against the right margin, so the dot leader spans the gap.
    static int TextWidth(SectionProperties? section)
    {
        var size = section?.GetFirstChild<PageSize>();
        var margin = section?.GetFirstChild<PageMargin>();
        if (size?.Width?.Value is not { } width)
        {
            // A4 less the margins Word defaults to.
            return 9026;
        }

        var left = margin?.Left?.Value ?? 1440;
        var right = margin?.Right?.Value ?? 1440;
        return (int)width - (int)left - (int)right;
    }

    // What Word's default table style puts either side of a cell's text.
    const int DefaultCellMargin = 108;

    // Tab stops in a cell are measured from the cell's own text edge, so a stop at the section's
    // width is past the cell's right edge and the page number is pushed out of the cell or onto a
    // line of its own. Narrowed one table at a time from the outside in, because a nested table is
    // sized against the cell holding it.
    static int WidthInCells(Paragraph paragraph, int width)
    {
        foreach (var cell in paragraph.Ancestors<TableCell>().Reverse())
        {
            width = InnerWidth(cell, width);
        }

        return width;
    }

    // The cell's share of its table, less the cell margins. Column widths are read as proportions of
    // the table rather than as measurements, which is what they are in a table sized by percentage
    // and comes to the same thing in one sized absolutely. A table with no column widths is on
    // autofit, where the columns follow their content and only Word knows where they end up: each
    // is taken to get an even share, which is exact for the single-column table a table of contents
    // is usually boxed in.
    static int InnerWidth(TableCell cell, int available)
    {
        var row = cell.Ancestors<TableRow>().FirstOrDefault();
        var table = row?.Ancestors<Table>().FirstOrDefault();
        if (row == null ||
            table == null)
        {
            return available;
        }

        var index = 0;
        var rowSpan = 0;
        foreach (var sibling in row.Elements<TableCell>())
        {
            if (sibling == cell)
            {
                index = rowSpan;
            }

            rowSpan += Span(sibling);
        }

        var span = Span(cell);
        var columns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Select(ColumnWidth).ToList() ?? [];
        var columnTotal = columns.Sum(_ => (long)_);
        var measured = columns.Count >= index + span &&
                       columns.All(_ => _ > 0);
        var properties = table.GetFirstChild<TableProperties>();

        long share;
        if (measured)
        {
            share = TableWidth(properties, available, columnTotal) *
                    columns.Skip(index).Take(span).Sum(_ => (long)_) /
                    columnTotal;
        }
        else
        {
            share = TableWidth(properties, available, 0) *
                    span /
                    Math.Max(Math.Max(columns.Count, rowSpan), 1);
        }

        var cellMargins = cell.TableCellProperties?.TableCellMargin;
        var tableMargins = properties?.TableCellMarginDefault;
        var inner = share -
                    Margin(cellMargins, tableMargins, "left", "start") -
                    Margin(cellMargins, tableMargins, "right", "end");
        if (inner <= 0)
        {
            return available;
        }

        return (int)inner;
    }

    static long TableWidth(TableProperties? properties, int available, long columnTotal)
    {
        var width = properties?.TableWidth;
        var type = width?.Type?.Value;
        if (int.TryParse(width?.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) &&
            value > 0)
        {
            if (type == TableWidthUnitValues.Pct)
            {
                // Fiftieths of a percent of what the table sits in.
                return (long)available * value / 5000;
            }

            if (type == TableWidthUnitValues.Dxa)
            {
                return value;
            }
        }

        if (columnTotal > 0)
        {
            return columnTotal;
        }

        return Math.Max(available - (properties?.TableIndentation?.Width?.Value ?? 0), 1);
    }

    static int Span(TableCell cell) =>
        Math.Max(cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1, 1);

    static int ColumnWidth(GridColumn column)
    {
        if (int.TryParse(column.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var width))
        {
            return width;
        }

        return 0;
    }

    // A cell's own margin wins over its table's, side by side. Each side has two spellings, and the
    // elements carrying them differ between the cell's list and the table's, so they are matched by
    // name rather than by type.
    static int Margin(OpenXmlElement? cell, OpenXmlElement? table, string name, string alias) =>
        Margin(cell, name, alias) ??
        Margin(table, name, alias) ??
        DefaultCellMargin;

    static int? Margin(OpenXmlElement? margins, string name, string alias)
    {
        var side = margins?.ChildElements.FirstOrDefault(_ => _.LocalName == name || _.LocalName == alias);
        if (side == null)
        {
            return null;
        }

        var width = side.GetAttributes().FirstOrDefault(_ => _.LocalName == "w").Value;
        if (int.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }

        return null;
    }

    [GeneratedRegex(
        """
        \\o\s+"\d+-(\d+)"
        """)]
    private static partial Regex ParseMaxLevelRegex();
}
