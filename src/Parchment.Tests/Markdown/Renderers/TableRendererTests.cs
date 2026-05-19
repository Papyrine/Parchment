using MarkdigTable = Markdig.Extensions.Tables.Table;

public class TableRendererTests
{
    public class EmptyModel;

    [Test]
    public async Task PipeTableEmitsTableWithGridRowsAndHeaderFormatting()
    {
        const string md =
            """
            | A | B |
            |---|---|
            | 1 | 2 |
            | 3 | 4 |
            """;

        var tableBlock = RendererHarness.FirstBlock<MarkdigTable>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(tableBlock);

        var table = (Table) renderer.Drain().Single();

        var grid = table.GetFirstChild<TableGrid>()!;
        await Assert.That(grid.Elements<GridColumn>().Count()).IsEqualTo(2);

        var rows = table.Elements<TableRow>().ToList();
        await Assert.That(rows.Count).IsEqualTo(3);

        var headerCells = rows[0].Elements<TableCell>().ToList();
        await Assert.That(headerCells.Count).IsEqualTo(2);
        var headerParagraph = headerCells[0].GetFirstChild<Paragraph>()!;
        await Assert.That(headerParagraph.ParagraphProperties!.GetFirstChild<Justification>()!.Val?.Value)
            .IsEqualTo(JustificationValues.Center);
        var headerRun = headerParagraph.GetFirstChild<Run>()!;
        await Assert.That(headerRun.RunProperties!.GetFirstChild<Bold>()).IsNotNull();

        var bodyCellRun = rows[1].Elements<TableCell>().First()
            .GetFirstChild<Paragraph>()!
            .GetFirstChild<Run>()!;
        await Assert.That(bodyCellRun.RunProperties?.GetFirstChild<Bold>()).IsNull();

        await VerifyDocument(md);
    }

    [Test]
    public Task PipeTableSetsFullWidthOnTableProperties()
    {
        const string md =
            """
            | A | B |
            |---|---|
            | 1 | 2 |
            """;

        var tableBlock = RendererHarness.FirstBlock<MarkdigTable>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(tableBlock);

        var table = (Table) renderer.Drain().Single();
        return AssertFullWidth(table);
    }

    [Test]
    public async Task TableNestedInBlockQuoteIsIndented()
    {
        const string md =
            """
            > | A | B |
            > |---|---|
            > | 1 | 2 |
            """;

        var quote = RendererHarness.FirstBlock<QuoteBlock>(md);
        await Assert.That(quote.Descendants<MarkdigTable>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(quote);

        var table = renderer.Drain().OfType<Table>().Single();
        await AssertIndented(table, expectedIndent: 720);
        await Assert.That(table.Elements<TableRow>().Count()).IsEqualTo(2);

        await VerifyDocument(md);
    }

    [Test]
    public async Task TableNestedInListItemIsIndented()
    {
        const string md =
            """
            - item

                | A | B |
                |---|---|
                | 1 | 2 |
            """;

        var list = RendererHarness.FirstBlock<ListBlock>(md);
        await Assert.That(list.Descendants<MarkdigTable>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(list);

        var table = renderer.Drain().OfType<Table>().Single();
        await AssertIndented(table, expectedIndent: 480);
        await Assert.That(table.Elements<TableRow>().Count()).IsEqualTo(2);

        await VerifyDocument(md);
    }

    [Test]
    public async Task TableNestedInNestedBlockQuoteAccumulatesIndent()
    {
        const string md =
            """
            > > | A | B |
            > > |---|---|
            > > | 1 | 2 |
            """;

        var outerQuote = RendererHarness.FirstBlock<QuoteBlock>(md);
        await Assert.That(outerQuote.Descendants<MarkdigTable>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(outerQuote);

        var table = renderer.Drain().OfType<Table>().Single();
        await AssertIndented(table, expectedIndent: 1440);

        await VerifyDocument(md);
    }

    static async Task AssertFullWidth(Table table)
    {
        var properties = table.GetFirstChild<TableProperties>()!;
        var width = properties.GetFirstChild<TableWidth>()!;
        await Assert.That(width.Width?.Value).IsEqualTo("5000");
        await Assert.That(width.Type?.Value).IsEqualTo(TableWidthUnitValues.Pct);
        await Assert.That(properties.GetFirstChild<TableIndentation>()).IsNull();
    }

    static async Task AssertIndented(Table table, int expectedIndent)
    {
        var properties = table.GetFirstChild<TableProperties>()!;
        var width = properties.GetFirstChild<TableWidth>()!;
        await Assert.That(width.Type?.Value).IsEqualTo(TableWidthUnitValues.Auto);
        var indent = properties.GetFirstChild<TableIndentation>()!;
        await Assert.That(indent.Width?.Value).IsEqualTo(expectedIndent);
        await Assert.That(indent.Type?.Value).IsEqualTo(TableWidthUnitValues.Dxa);
    }

    [Test]
    public async Task GridTableEmitsTableWithCorrectStructure()
    {
        const string md =
            """
            +---+---+
            | A | B |
            +===+===+
            | 1 | 2 |
            +---+---+
            | 3 | 4 |
            +---+---+
            """;

        var tableBlock = RendererHarness.FirstBlock<MarkdigTable>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(tableBlock);

        var table = (Table) renderer.Drain().Single();

        var grid = table.GetFirstChild<TableGrid>()!;
        await Assert.That(grid.Elements<GridColumn>().Count()).IsEqualTo(2);

        var rows = table.Elements<TableRow>().ToList();
        await Assert.That(rows.Count).IsEqualTo(3);

        var headerCells = rows[0].Elements<TableCell>().ToList();
        await Assert.That(headerCells.Count).IsEqualTo(2);
        var headerParagraph = headerCells[0].GetFirstChild<Paragraph>()!;
        await Assert.That(headerParagraph.ParagraphProperties!.GetFirstChild<Justification>()!.Val?.Value)
            .IsEqualTo(JustificationValues.Center);
        var headerRun = headerParagraph.GetFirstChild<Run>()!;
        await Assert.That(headerRun.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
        await Assert.That(headerRun.GetFirstChild<Text>()!.Text).IsEqualTo("A");

        var bodyCell = rows[1].Elements<TableCell>().First()
            .GetFirstChild<Paragraph>()!
            .GetFirstChild<Run>()!;
        await Assert.That(bodyCell.GetFirstChild<Text>()!.Text).IsEqualTo("1");
        await Assert.That(bodyCell.RunProperties?.GetFirstChild<Bold>()).IsNull();

        await VerifyDocument(md);
    }

    [Test]
    public async Task PipeTableHonorsColumnAlignment()
    {
        const string md =
            """
            | Left | Center | Right | Default |
            | :- | :-: | -: | -|
            | a    | b      | c     | d       |
            | e    | f      | g     | h       |
            """;

        var tableBlock = RendererHarness.FirstBlock<MarkdigTable>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(tableBlock);

        var table = (Table) renderer.Drain().Single();
        var rows = table.Elements<TableRow>().ToList();

        JustificationValues?[] expectedHeader =
        [
            JustificationValues.Left,
            JustificationValues.Center,
            JustificationValues.Right,
            JustificationValues.Center
        ];
        JustificationValues?[] expectedBody =
        [
            JustificationValues.Left,
            JustificationValues.Center,
            JustificationValues.Right,
            null
        ];

        await AssertRowJustifications(rows[0], expectedHeader);
        await AssertRowJustifications(rows[1], expectedBody);
        await AssertRowJustifications(rows[2], expectedBody);

        await VerifyDocument(md);
    }

    static async Task AssertRowJustifications(TableRow row, JustificationValues?[] expected)
    {
        var cells = row.Elements<TableCell>().ToList();
        for (var i = 0; i < expected.Length; i++)
        {
            var paragraph = cells[i].GetFirstChild<Paragraph>()!;
            var actual = paragraph.ParagraphProperties?.GetFirstChild<Justification>()?.Val?.Value;
            if (expected[i] is { } value)
            {
                await Assert.That(actual).IsEqualTo(value);
            }
            else
            {
                await Assert.That(actual).IsNull();
            }
        }
    }

    [Test]
    public async Task PipeTableHonorsColumnWidthsFromDashCounts()
    {
        // Separator dash counts 6 / 20 / 6 → widths 18.75% / 62.5% / 18.75%.
        // At a 9000 dxa budget that maps to 1688 / 5625 / 1688.
        const string md =
            """
            | Name | Description           | Count |
            |------|--------------------|------|
            | A    | Short                 | 1     |
            | BB   | A longer description  | 22    |
            """;

        var tableBlock = RendererHarness.FirstBlock<MarkdigTable>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(tableBlock);

        var table = (Table) renderer.Drain().Single();

        var properties = table.GetFirstChild<TableProperties>()!;
        await Assert.That(properties.GetFirstChild<TableLayout>()?.Type?.Value)
            .IsEqualTo(TableLayoutValues.Fixed);

        var gridCols = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
        await Assert.That(gridCols.Count).IsEqualTo(3);
        await Assert.That(gridCols[0].Width?.Value).IsEqualTo("1688");
        await Assert.That(gridCols[1].Width?.Value).IsEqualTo("5625");
        await Assert.That(gridCols[2].Width?.Value).IsEqualTo("1688");

        var firstBodyRow = table.Elements<TableRow>().ElementAt(1);
        var cells = firstBodyRow.Elements<TableCell>().ToList();
        for (var i = 0; i < cells.Count; i++)
        {
            var cellWidth = cells[i].TableCellProperties?.GetFirstChild<TableCellWidth>();
            await Assert.That(cellWidth?.Width?.Value).IsEqualTo(gridCols[i].Width?.Value);
            await Assert.That(cellWidth?.Type?.Value).IsEqualTo(TableWidthUnitValues.Dxa);
        }

        await VerifyDocument(md);
    }

    [Test]
    public async Task PipeTableWithUniformSeparatorPreservesEqualColumns()
    {
        const string md =
            """
            | A | B | C |
            |---|---|---|
            | 1 | 2 | 3 |
            """;

        var tableBlock = RendererHarness.FirstBlock<MarkdigTable>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(tableBlock);

        var table = (Table) renderer.Drain().Single();
        var gridCols = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
        await Assert.That(gridCols.Count).IsEqualTo(3);
        await Assert.That(gridCols[0].Width?.Value).IsEqualTo(gridCols[1].Width?.Value);
        await Assert.That(gridCols[1].Width?.Value).IsEqualTo(gridCols[2].Width?.Value);
    }

    [Test]
    public async Task IndentedTableSkipsExplicitColumnWidths()
    {
        // Uneven dashes would normally inflate the inner Description column. Inside a blockquote
        // the table is auto-sized and explicit dxa widths would override the indent, so the
        // renderer skips the width emission for indented tables.
        const string md =
            """
            > | Name | Description           |
            > |------|--------------------|
            > | A    | Short                 |
            """;

        var quote = RendererHarness.FirstBlock<QuoteBlock>(md);
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(quote);

        var table = renderer.Drain().OfType<Table>().Single();
        var properties = table.GetFirstChild<TableProperties>()!;
        await Assert.That(properties.GetFirstChild<TableLayout>()).IsNull();

        var gridCols = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
        foreach (var col in gridCols)
        {
            await Assert.That(col.Width).IsNull();
        }
    }

    static async Task VerifyDocument(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("table", markdown, styleSource);
        using var stream = new MemoryStream();
        await store.Render("table", new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
