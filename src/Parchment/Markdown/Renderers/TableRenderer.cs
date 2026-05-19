class TableRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, Markdig.Extensions.Tables.Table>
{
    // Approximate page-content width budget in dxa (twentieths of a point). When
    // ColumnDefinitions carry width percentages, per-column dxa values are computed
    // proportionally from this budget. For Pct-width tables (the non-indented case)
    // Word treats these as ratios; for indented tables they are absolute.
    const int GridWidthBudgetDxa = 9000;

    protected override void Write(OpenXmlMarkdownRenderer renderer, Markdig.Extensions.Tables.Table tableBlock)
    {
        var columns = tableBlock.ColumnDefinitions;
        // Indented tables use Auto width and size to content; absolute dxa column widths would
        // override that and stretch the table to the full budget. Tables flagged by
        // SkipColumnWidths had aligned pipes across header/separator/body in the source,
        // signalling readability padding rather than custom widths.
        var columnWidths = renderer.CurrentIndent > 0 || renderer.SkipColumnWidths.Contains(tableBlock)
            ? null
            : ComputeColumnWidths(columns);

        var table = new Table();
        table.Append(BuildTableProperties(renderer.CurrentIndent, columnWidths is not null));
        table.Append(BuildTableGrid(tableBlock, columnWidths));

        foreach (var child in tableBlock)
        {
            if (child is Markdig.Extensions.Tables.TableRow row)
            {
                table.Append(BuildRow(renderer, row, columns, columnWidths));
            }
        }

        renderer.AddBlock(table);
    }

    static int[]? ComputeColumnWidths(IList<Markdig.Extensions.Tables.TableColumnDefinition> columns)
    {
        if (columns.Count == 0)
        {
            return null;
        }

        float totalPct = 0;
        var first = columns[0].Width;
        var allEqual = true;
        foreach (var column in columns)
        {
            totalPct += column.Width;
            if (column.Width != first)
            {
                allEqual = false;
            }
        }

        // No width hints from Markdig (totalPct == 0) or every column has the same width —
        // either way the default Word auto-distribution yields the same layout, so skip the
        // explicit emission to keep the docx output minimal. Source-aligned pipe tables are
        // filtered upstream via OpenXmlMarkdownRenderer.SkipColumnWidths.
        if (totalPct <= 0 || allEqual)
        {
            return null;
        }

        var widths = new int[columns.Count];
        for (var i = 0; i < columns.Count; i++)
        {
            var pct = columns[i].Width;
            widths[i] = pct > 0
                ? Math.Max(1, (int) Math.Round(GridWidthBudgetDxa * pct / totalPct))
                : 0;
        }

        return widths;
    }

    static TableProperties BuildTableProperties(int indent, bool hasColumnWidths)
    {
        var width = indent > 0
            ? new TableWidth { Type = TableWidthUnitValues.Auto }
            : new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
        var properties = new TableProperties(width);
        if (indent > 0)
        {
            properties.Append(
                new TableIndentation
                {
                    Width = indent,
                    Type = TableWidthUnitValues.Dxa
                });
        }

        if (hasColumnWidths)
        {
            properties.Append(new TableLayout { Type = TableLayoutValues.Fixed });
        }

        properties.Append(BuildBorders());
        properties.Append(BuildCellMargins());
        return properties;
    }

    static TableBorders BuildBorders() =>
        new(
            new TopBorder
            {
                Val = BorderValues.Single,
                Size = 4
            },
            new BottomBorder
            {
                Val = BorderValues.Single,
                Size = 4
            },
            new LeftBorder
            {
                Val = BorderValues.Single,
                Size = 4
            },
            new RightBorder
            {
                Val = BorderValues.Single,
                Size = 4
            },
            new InsideHorizontalBorder
            {
                Val = BorderValues.Single,
                Size = 4
            },
            new InsideVerticalBorder
            {
                Val = BorderValues.Single,
                Size = 4
            });

    static TableCellMarginDefault BuildCellMargins() =>
        new(
            new TopMargin
            {
                Width = "0",
                Type = TableWidthUnitValues.Dxa
            },
            new StartMargin
            {
                Width = "108",
                Type = TableWidthUnitValues.Dxa
            },
            new BottomMargin
            {
                Width = "0",
                Type = TableWidthUnitValues.Dxa
            },
            new EndMargin
            {
                Width = "108",
                Type = TableWidthUnitValues.Dxa
            });

    static TableGrid BuildTableGrid(Markdig.Extensions.Tables.Table table, int[]? columnWidths)
    {
        var grid = new TableGrid();
        var columns = table.ColumnDefinitions.Count;
        for (var i = 0; i < columns; i++)
        {
            var gridColumn = new GridColumn();
            if (columnWidths is not null && columnWidths[i] > 0)
            {
                gridColumn.Width = columnWidths[i].ToString(CultureInfo.InvariantCulture);
            }

            grid.Append(gridColumn);
        }

        return grid;
    }

    static TableRow BuildRow(
        OpenXmlMarkdownRenderer renderer,
        Markdig.Extensions.Tables.TableRow row,
        IList<Markdig.Extensions.Tables.TableColumnDefinition> columns,
        int[]? columnWidths)
    {
        var tableRow = new TableRow();
        var index = 0;
        foreach (var cell in row.OfType<Markdig.Extensions.Tables.TableCell>())
        {
            var columnIndex = cell.ColumnIndex >= 0 ? cell.ColumnIndex : index;
            var alignment = columnIndex < columns.Count ? columns[columnIndex].Alignment : null;
            var width = columnWidths is not null && columnIndex < columnWidths.Length
                ? columnWidths[columnIndex]
                : 0;
            tableRow.Append(BuildCell(renderer, cell, row.IsHeader, alignment, width));
            index += cell.ColumnSpan > 0 ? cell.ColumnSpan : 1;
        }

        return tableRow;
    }

    static TableCell BuildCell(
        OpenXmlMarkdownRenderer renderer,
        Markdig.Extensions.Tables.TableCell cell,
        bool isHeader,
        Markdig.Extensions.Tables.TableColumnAlign? alignment,
        int widthDxa)
    {
        // Fast path: data-table cells are overwhelmingly a single ParagraphBlock containing one
        // LiteralInline (plain text). Skip the PushContainer / Render / PopContainer dance and
        // synthesize the OpenXml subtree directly. Falls through to the general path for any
        // structural variant — emphasis, links, multiple paragraphs, embedded HTML, etc.
        if (TryBuildPlainCell(cell, isHeader, alignment, widthDxa) is { } fast)
        {
            return fast;
        }

        var tableCell = new TableCell();
        ApplyCellWidth(tableCell, widthDxa);
        renderer.PushContainer();
        foreach (var child in cell)
        {
            renderer.Render(child);
        }

        var state = renderer.PopContainer();
        if (state.CurrentRuns.Count > 0)
        {
            var paragraph = new Paragraph();
            foreach (var run in state.CurrentRuns)
            {
                paragraph.Append(run);
            }

            state.Blocks.Add(paragraph);
        }

        if (state.Blocks.Count == 0)
        {
            state.Blocks.Add(new Paragraph());
        }

        foreach (var block in state.Blocks)
        {
            if (block is Paragraph p)
            {
                ApplyCellFormatting(p, isHeader, alignment);
            }

            tableCell.Append(block);
        }

        renderer.ReleaseContainer(state);
        return tableCell;
    }

    static TableCell? TryBuildPlainCell(
        Markdig.Extensions.Tables.TableCell cell,
        bool isHeader,
        Markdig.Extensions.Tables.TableColumnAlign? alignment,
        int widthDxa)
    {
        if (cell is not [ParagraphBlock paragraphBlock])
        {
            return null;
        }

        var inline = paragraphBlock.Inline;
        if (inline?.FirstChild is not LiteralInline {NextSibling: null} literal)
        {
            return null;
        }

        var paragraph = new Paragraph();
        var content = literal.Content.AsSpan();
        if (content.Length > 0)
        {
            var run = new Run(
                new Text(XmlCharSanitizer.Strip(content).ToString())
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            if (isHeader)
            {
                run.RunProperties = new();
                run.RunProperties.Append(new Bold());
            }

            paragraph.Append(run);
        }

        var justification = ResolveJustification(isHeader, alignment);
        if (justification is not null)
        {
            paragraph.ParagraphProperties = new();
            paragraph.ParagraphProperties.Append(new Justification { Val = justification });
        }

        var tableCell = new TableCell(paragraph);
        ApplyCellWidth(tableCell, widthDxa);
        return tableCell;
    }

    static void ApplyCellWidth(TableCell tableCell, int widthDxa)
    {
        if (widthDxa <= 0)
        {
            return;
        }

        tableCell.TableCellProperties = new(
            new TableCellWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = widthDxa.ToString(CultureInfo.InvariantCulture)
            });
    }

    static void ApplyCellFormatting(
        Paragraph paragraph,
        bool isHeader,
        Markdig.Extensions.Tables.TableColumnAlign? alignment)
    {
        var justification = ResolveJustification(isHeader, alignment);
        if (justification is not null)
        {
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.Append(new Justification { Val = justification });
        }

        if (!isHeader)
        {
            return;
        }

        foreach (var run in paragraph.Descendants<Run>())
        {
            run.RunProperties ??= new();
            run.RunProperties.Append(new Bold());
        }
    }

    static JustificationValues? ResolveJustification(
        bool isHeader,
        Markdig.Extensions.Tables.TableColumnAlign? alignment) =>
        alignment switch
        {
            Markdig.Extensions.Tables.TableColumnAlign.Left => JustificationValues.Left,
            Markdig.Extensions.Tables.TableColumnAlign.Center => JustificationValues.Center,
            Markdig.Extensions.Tables.TableColumnAlign.Right => JustificationValues.Right,
            _ => isHeader ? JustificationValues.Center : null
        };
}
