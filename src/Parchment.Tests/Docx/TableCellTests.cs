// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// Word refuses to open a document holding a <c>w:tc</c> that does not end with a <c>w:p</c>, while
/// the schema — and so <c>OpenXmlValidator</c> — accepts one. Each test leaves a cell with nothing
/// but the token under test, so whatever the token removes or produces is all the cell has.
/// </summary>
public partial class TableCellTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class CellModel
    {
        public string Name { get; init; } = "Alice";
        public bool Flag { get; init; }
        public IReadOnlyList<Line> Lines { get; init; } = [];
        public string[] Tags { get; init; } = [];

        [Html]
        public string Html { get; init; } = "";

        [Markdown]
        public string Markdown { get; init; } = "";

        [ExcelsiorTable]
        public IReadOnlyList<Line>? Rows { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
    }

    [Test]
    public async Task FalseIfLeavesCellWithParagraph()
    {
        using var doc = await Render(new(), "{% if Flag %}", "{{ Name }}", "{% endif %}");

        await AssertCellsEndWithParagraph(doc);
        await Assert.That(Cell(doc).InnerText).IsEqualTo("");
    }

    [Test]
    public async Task FalseIfWithoutElseBranchTakenLeavesCellWithParagraph()
    {
        using var doc = await Render(
            new(),
            "{% if Flag %}",
            "{{ Name }}",
            "{% elsif Flag %}",
            "other",
            "{% endif %}");

        await AssertCellsEndWithParagraph(doc);
        await Assert.That(Cell(doc).InnerText).IsEqualTo("");
    }

    [Test]
    public async Task EmptyLoopLeavesCellWithParagraph()
    {
        using var doc = await Render(new(), "{% for line in Lines %}", "{{ line.Name }}", "{% endfor %}");

        await AssertCellsEndWithParagraph(doc);
        await Assert.That(Cell(doc).InnerText).IsEqualTo("");
    }

    [Test]
    public async Task EmptyHtmlLeavesCellWithParagraph()
    {
        using var doc = await Render(new(), "{{ Html }}");

        await AssertCellsEndWithParagraph(doc);
    }

    [Test]
    public async Task EmptyMarkdownLeavesCellWithParagraph()
    {
        using var doc = await Render(new(), "{{ Markdown }}");

        await AssertCellsEndWithParagraph(doc);
    }

    [Test]
    public async Task EmptyStringListLeavesCellWithParagraph()
    {
        using var doc = await Render(new(), "{{ Tags }}");

        await AssertCellsEndWithParagraph(doc);
    }

    [Test]
    public async Task NullExcelsiorTableLeavesCellWithParagraph()
    {
        using var doc = await Render(new(), "{{ Rows }}");

        await AssertCellsEndWithParagraph(doc);
    }

    [Test]
    public async Task NestedExcelsiorTableIsFollowedByParagraph()
    {
        using var doc = await Render(
            new()
            {
                Rows = [new() { Name = "a" }]
            },
            "{{ Rows }}");

        await AssertCellsEndWithParagraph(doc);

        // The nested table is still there, and still ahead of the paragraph that closes the cell.
        var children = Cell(doc).ChildElements;
        await Assert.That(children[^2]).IsTypeOf<Table>();
        await Assert.That(children[^1].InnerText).IsEqualTo("");
    }

    [Test]
    public async Task CellThatKeepsItsParagraphIsUntouched()
    {
        using var doc = await Render(
            new()
            {
                Flag = true
            },
            "{% if Flag %}",
            "{{ Name }}",
            "{% endif %}");

        await AssertCellsEndWithParagraph(doc);
        var paragraph = Cell(doc).ChildElements.Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("Alice");
    }

    [Test]
    public async Task HeaderCellIsRepaired()
    {
        using var template = BuildTemplate("body");
        using (var source = WordprocessingDocument.Open(template, true))
        {
            var mainPart = source.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new(OneCell("{% if Flag %}", "{{ Name }}", "{% endif %}"), new Paragraph());
            mainPart.Document!.Body!.GetFirstChild<SectionProperties>()!.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = mainPart.GetIdOfPart(headerPart)
                });
        }

        template.Position = 0;
        using var doc = await Render(template, new());

        var cell = doc.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<TableCell>().Single();
        await Assert.That(cell.LastChild).IsTypeOf<Paragraph>();
    }

    [Test]
    public async Task MarkdownCellWithEmptyHtmlHasParagraph()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<CellModel>(
            """
            | A |
            |---|
            | {{ Html }} |
            """);

        using var doc = await Render(store, new());

        await AssertCellsEndWithParagraph(doc);
    }

    [Test]
    public async Task MarkdownCellHoldingHtmlTableIsFollowedByParagraph()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<CellModel>(
            """
            | A |
            |---|
            | {{ Html }} |
            """);

        using var doc = await Render(
            store,
            new()
            {
                Html = "<table><tr><td>x</td></tr></table>"
            });

        await AssertCellsEndWithParagraph(doc);
        var outer = doc.MainDocumentPart!.Document!.Body!.GetFirstChild<Table>()!;
        await Assert.That(outer.Descendants<Table>().Count()).IsEqualTo(1);
    }

    static TableCell Cell(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!.GetFirstChild<Table>()!.Descendants<TableCell>().First();

    static async Task AssertCellsEndWithParagraph(WordprocessingDocument doc)
    {
        var cells = doc.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().ToList();
        await Assert.That(cells.Count).IsGreaterThan(0);
        foreach (var cell in cells)
        {
            await Assert.That(cell.LastChild).IsTypeOf<Paragraph>();
        }
    }

    static Task<WordprocessingDocument> Render(CellModel model, params string[] cellParagraphs)
    {
        using var template = BuildTemplate(cellParagraphs);
        return Render(template, model);
    }

    static Task<WordprocessingDocument> Render(Stream template, CellModel model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<CellModel>(template);
        return Render(store, model);
    }

    static async Task<WordprocessingDocument> Render(TemplateStore store, CellModel model)
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    /// <summary>
    /// A body holding a one-cell table whose cell has one paragraph per entry, then a paragraph.
    /// </summary>
    static MemoryStream BuildTemplate(params string[] cellParagraphs)
    {
        var stream = DocxTemplateBuilder.Build("after");
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            doc.MainDocumentPart!.Document!.Body!.PrependChild(OneCell(cellParagraphs));
        }

        stream.Position = 0;
        return stream;
    }

    static Table OneCell(params string[] cellParagraphs)
    {
        var cell = new TableCell();
        foreach (var text in cellParagraphs)
        {
            cell.Append(new Paragraph(new Run(new Text(text))));
        }

        return new(
            new TableProperties(
                new TableWidth
                {
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(new GridColumn()),
            new TableRow(cell));
    }
}
