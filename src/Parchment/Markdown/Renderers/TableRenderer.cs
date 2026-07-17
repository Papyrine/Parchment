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

        // {.StyleName} has to lead the table, on its own line, matching how it is written for a
        // heading or paragraph. A trailing one is not an option: directly after the last row it
        // stops the table parsing at all, and after a blank line it binds to nothing.
        var styleId = MarkdownStyle.Resolve(tableBlock);

        var table = new Table();
        table.Append(BuildTableProperties(renderer.CurrentIndent, columnWidths is not null, styleId));
        table.Append(BuildTableGrid(tableBlock, columnWidths));

        foreach (var child in tableBlock)
        {
            if (child is Markdig.Extensions.Tables.TableRow row)
            {
                table.Append(BuildRow(renderer, row, columns, columnWidths, styleId is not null));
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

        // Skip when Markdig gave no width hints (totalPct == 0), and when every column is the same
        // width. Uniform separators are the conventional way to write a table (`| --- | --- |`)
        // rather than a request for equal columns, so they are treated as "no opinion" and the
        // table is left on Word's autofit, which sizes columns to their content. Source-aligned
        // pipe tables — the other conventional style, where the dashes are padded for readability —
        // are filtered upstream via OpenXmlMarkdownRenderer.SkipColumnWidths.
        //
        // The cost is that genuinely equal columns cannot be stated: equal dash counts are
        // indistinguishable from the conventional separator. Expressing that needs explicit width
        // syntax, not a dash-count heuristic.
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

    static TableProperties BuildTableProperties(int indent, bool hasColumnWidths, string? styleId)
    {
        var properties = new TableProperties();

        // tblStyle leads the tblPr sequence.
        if (styleId != null)
        {
            properties.Append(
                new TableStyle
                {
                    Val = styleId
                });
        }

        properties.Append(
            indent > 0
                ? new TableWidth { Type = TableWidthUnitValues.Auto }
                : new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });

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

        // Direct formatting beats a table style in Word, so emitting the default borders and cell
        // margins alongside a tblStyle would silently override the style the caller asked for.
        // Without a style they are the only thing making the table legible, so they stay.
        if (styleId == null)
        {
            properties.Append(BuildBorders());
            properties.Append(BuildCellMargins());
        }

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
        int[]? columnWidths,
        bool hasTableStyle)
    {
        var tableRow = new TableRow();

        // Repeat the header on every page a table spans. Without this a table broken across a page
        // break loses its header entirely on the later pages.
        if (row.IsHeader)
        {
            tableRow.Append(new TableRowProperties(new TableHeader()));
        }

        var index = 0;
        foreach (var cell in row.OfType<Markdig.Extensions.Tables.TableCell>())
        {
            var columnIndex = cell.ColumnIndex >= 0 ? cell.ColumnIndex : index;
            var alignment = columnIndex < columns.Count ? columns[columnIndex].Alignment : null;
            var span = cell.ColumnSpan > 0 ? cell.ColumnSpan : 1;

            // A spanning cell covers its own column plus the ones it swallows, so it gets their
            // combined width.
            var width = 0;
            if (columnWidths is not null)
            {
                for (var i = columnIndex; i < columnIndex + span && i < columnWidths.Length; i++)
                {
                    width += columnWidths[i];
                }
            }

            // A table style supplies its own header formatting through conditional formatting
            // (firstRow), so the hardcoded bold and centring stand down rather than override it.
            var isHeader = row.IsHeader && !hasTableStyle;
            tableRow.Append(BuildCell(renderer, cell, isHeader, alignment, width, span));
            index += span;
        }

        return tableRow;
    }

    static TableCell BuildCell(
        OpenXmlMarkdownRenderer renderer,
        Markdig.Extensions.Tables.TableCell cell,
        bool isHeader,
        Markdig.Extensions.Tables.TableColumnAlign? alignment,
        int widthDxa,
        int columnSpan)
    {
        // Fast path: data-table cells are overwhelmingly a single ParagraphBlock containing one
        // LiteralInline (plain text). Skip the PushContainer / Render / PopContainer dance and
        // synthesize the OpenXml subtree directly. Falls through to the general path for any
        // structural variant — emphasis, links, multiple paragraphs, embedded HTML, etc.
        if (TryBuildPlainCell(cell, isHeader, alignment, widthDxa, columnSpan) is { } fast)
        {
            return fast;
        }

        var tableCell = new TableCell();
        ApplyCellProperties(tableCell, widthDxa, columnSpan);
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
        int widthDxa,
        int columnSpan)
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
        ApplyCellProperties(tableCell, widthDxa, columnSpan);
        return tableCell;
    }

    // Grid-table colspan was read only to advance the column cursor and never emitted, so a
    // spanning row ended up with fewer <w:tc> than the grid had <w:gridCol> — a silently ragged
    // table rather than an error.
    static void ApplyCellProperties(TableCell tableCell, int widthDxa, int columnSpan)
    {
        if (widthDxa <= 0 &&
            columnSpan <= 1)
        {
            return;
        }

        // tcW precedes gridSpan in the tcPr schema sequence.
        var properties = new TableCellProperties();
        if (widthDxa > 0)
        {
            properties.Append(
                new TableCellWidth
                {
                    Type = TableWidthUnitValues.Dxa,
                    Width = widthDxa.ToString(CultureInfo.InvariantCulture)
                });
        }

        if (columnSpan > 1)
        {
            properties.Append(
                new GridSpan
                {
                    Val = columnSpan
                });
        }

        tableCell.TableCellProperties = properties;
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
