class TableRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, Markdig.Extensions.Tables.Table>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, Markdig.Extensions.Tables.Table tableBlock)
    {
        var table = new Table();
        table.Append(BuildTableProperties());
        table.Append(BuildTableGrid(tableBlock));

        foreach (var child in tableBlock)
        {
            if (child is Markdig.Extensions.Tables.TableRow row)
            {
                table.Append(BuildRow(renderer, row));
            }
        }

        renderer.AddBlock(table);
    }

    static TableProperties BuildTableProperties() =>
        new(
            new TableBorders(
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
                }),
            new TableCellMarginDefault(
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
                }));

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

    static TableRow BuildRow(OpenXmlMarkdownRenderer renderer, Markdig.Extensions.Tables.TableRow row)
    {
        var tableRow = new TableRow();
        foreach (var cell in row.OfType<Markdig.Extensions.Tables.TableCell>())
        {
            tableRow.Append(BuildCell(renderer, cell, row.IsHeader));
        }

        return tableRow;
    }

    static TableCell BuildCell(OpenXmlMarkdownRenderer renderer, Markdig.Extensions.Tables.TableCell cell, bool isHeader)
    {
        // Fast path: data-table cells are overwhelmingly a single ParagraphBlock containing one
        // LiteralInline (plain text). Skip the PushContainer / Render / PopContainer dance and
        // synthesize the OpenXml subtree directly. Falls through to the general path for any
        // structural variant — emphasis, links, multiple paragraphs, embedded HTML, etc.
        if (TryBuildPlainCell(cell, isHeader) is { } fast)
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
            if (block is Paragraph p && isHeader)
            {
                ApplyHeaderFormatting(p);
            }

            tableCell.Append(block);
        }

        renderer.ReleaseContainer(state);
        return tableCell;
    }

    static TableCell? TryBuildPlainCell(Markdig.Extensions.Tables.TableCell cell, bool isHeader)
    {
        if (cell.Count != 1 ||
            cell[0] is not ParagraphBlock paragraphBlock)
        {
            return null;
        }

        var inline = paragraphBlock.Inline;
        if (inline?.FirstChild is not LiteralInline literal ||
            literal.NextSibling != null)
        {
            return null;
        }

        var paragraph = new Paragraph();
        var text = literal.Content.ToString();
        if (text.Length > 0)
        {
            var run = new Run(
                new Text(XmlCharSanitizer.Strip(text))
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

        if (isHeader)
        {
            paragraph.ParagraphProperties = new();
            paragraph.ParagraphProperties.Append(
                new Justification
                {
                    Val = JustificationValues.Center
                });
        }

        return new TableCell(paragraph);
    }

    static void ApplyHeaderFormatting(Paragraph p)
    {
        p.ParagraphProperties ??= new();
        p.ParagraphProperties.Append(
            new Justification
            {
                Val = JustificationValues.Center
            });
        foreach (var run in p.Elements<Run>())
        {
            run.RunProperties ??= new();
            run.RunProperties.Append(new Bold());
        }
    }
}