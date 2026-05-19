class TableRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, Markdig.Extensions.Tables.Table>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, Markdig.Extensions.Tables.Table tableBlock)
    {
        var table = new Table();
        table.Append(BuildTableProperties(renderer.CurrentIndent));
        table.Append(BuildTableGrid(tableBlock));

        var columns = tableBlock.ColumnDefinitions;
        foreach (var child in tableBlock)
        {
            if (child is Markdig.Extensions.Tables.TableRow row)
            {
                table.Append(BuildRow(renderer, row, columns));
            }
        }

        renderer.AddBlock(table);
    }

    static TableProperties BuildTableProperties(int indent)
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

    static TableGrid BuildTableGrid(Markdig.Extensions.Tables.Table table)
    {
        var grid = new TableGrid();
        var columns = table.ColumnDefinitions.Count;
        for (var i = 0; i < columns; i++)
        {
            grid.Append(new GridColumn());
        }

        return grid;
    }

    static TableRow BuildRow(
        OpenXmlMarkdownRenderer renderer,
        Markdig.Extensions.Tables.TableRow row,
        IList<Markdig.Extensions.Tables.TableColumnDefinition> columns)
    {
        var tableRow = new TableRow();
        var index = 0;
        foreach (var cell in row.OfType<Markdig.Extensions.Tables.TableCell>())
        {
            var columnIndex = cell.ColumnIndex >= 0 ? cell.ColumnIndex : index;
            var alignment = columnIndex < columns.Count ? columns[columnIndex].Alignment : null;
            tableRow.Append(BuildCell(renderer, cell, row.IsHeader, alignment));
            index += cell.ColumnSpan > 0 ? cell.ColumnSpan : 1;
        }

        return tableRow;
    }

    static TableCell BuildCell(
        OpenXmlMarkdownRenderer renderer,
        Markdig.Extensions.Tables.TableCell cell,
        bool isHeader,
        Markdig.Extensions.Tables.TableColumnAlign? alignment)
    {
        // Fast path: data-table cells are overwhelmingly a single ParagraphBlock containing one
        // LiteralInline (plain text). Skip the PushContainer / Render / PopContainer dance and
        // synthesize the OpenXml subtree directly. Falls through to the general path for any
        // structural variant — emphasis, links, multiple paragraphs, embedded HTML, etc.
        if (TryBuildPlainCell(cell, isHeader, alignment) is { } fast)
        {
            return fast;
        }

        var tableCell = new TableCell();
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
        Markdig.Extensions.Tables.TableColumnAlign? alignment)
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

        return new(paragraph);
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

        foreach (var run in paragraph.Elements<Run>())
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
