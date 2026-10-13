using DocumentFormat.OpenXml.Validation;

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
        AddHeader(template, OneCell("{% if Flag %}", "{{ Name }}", "{% endif %}"));
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

    // A block-level content control is not a paragraph: Word reads through it, so a cell whose
    // control ends with a table ends with a table. The paragraph that closes the cell goes after
    // the control, as a child of the cell, and never into the control's own content.
    [Test]
    public async Task ControlHoldingExcelsiorTableIsFollowedByParagraph()
    {
        using var doc = await Render(
            new()
            {
                Rows = [new() { Name = "a" }]
            },
            Control(Paragraph("{{ Rows }}")));

        await AssertControlThenParagraph(doc);
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    [Test]
    public async Task ControlHoldingHtmlTableIsFollowedByParagraph()
    {
        using var doc = await Render(
            new()
            {
                Html = "<table><tr><td>x</td></tr></table>"
            },
            Control(Paragraph("{{ Html }}")));

        await AssertControlThenParagraph(doc);
    }

    [Test]
    public async Task ControlHoldingMarkdownTableIsFollowedByParagraph()
    {
        using var doc = await Render(
            new()
            {
                Markdown =
                    """
                    | A |
                    |---|
                    | x |
                    """
            },
            Control(Paragraph("{{ Markdown }}")));

        await AssertControlThenParagraph(doc);
    }

    [Test]
    public async Task ControlLeftHoldingOnlyTemplateTableIsFollowedByParagraph()
    {
        using var doc = await Render(
            new(),
            Control(OneCell("inner"), Paragraph("{% if Flag %}"), Paragraph("{{ Name }}"), Paragraph("{% endif %}")));

        await AssertControlThenParagraph(doc);
    }

    [Test]
    public async Task ControlAfterParagraphHoldingTableIsFollowedByParagraph()
    {
        using var doc = await Render(
            new()
            {
                Rows = [new() { Name = "a" }]
            },
            Paragraph("{{ Name }}"),
            Control(Paragraph("{{ Rows }}")));

        await AssertCellsEndWithParagraph(doc);
        var children = Cell(doc).ChildElements;
        await Assert.That(children.Count).IsEqualTo(3);
        await Assert.That(children[0].InnerText).IsEqualTo("Alice");
        await Assert.That(children[1]).IsTypeOf<SdtBlock>();
        await Assert.That(children[2].InnerText).IsEqualTo("");
    }

    [Test]
    public async Task ControlThatEndsWithParagraphIsUntouched()
    {
        using var doc = await Render(
            new()
            {
                Rows = [new() { Name = "a" }]
            },
            Control(Paragraph("{{ Rows }}"), Paragraph("{{ Name }}")));

        // The control carries the paragraph that closes the cell, which is how Word writes a cell
        // whose content is one control.
        var control = (SdtBlock) Cell(doc).ChildElements.Single();
        var content = control.SdtContentBlock!.ChildElements;
        await Assert.That(content.Count).IsEqualTo(2);
        await Assert.That(content[0]).IsTypeOf<Table>();
        await Assert.That(content[1].InnerText).IsEqualTo("Alice");
    }

    [Test]
    public async Task ControlEmptiedByFalseIfIsUntouched()
    {
        using var doc = await Render(
            new(),
            Control(Paragraph("{% if Flag %}"), Paragraph("{{ Name }}"), Paragraph("{% endif %}")));

        // Word opens a cell whose only block is a control with no content.
        var control = (SdtBlock) Cell(doc).ChildElements.Single();
        await Assert.That(control.SdtContentBlock!.HasChildren).IsFalse();
    }

    [Test]
    public async Task HeaderCellControlHoldingTableIsFollowedByParagraph()
    {
        using var template = BuildTemplate("body");
        AddHeader(template, OneCell(Control(Paragraph("{{ Html }}"))));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<table><tr><td>x</td></tr></table>"
            });

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await AssertControlThenParagraph(header.GetFirstChild<Table>()!);
    }

    // A cell the template itself left open is closed at registration, so in a part no render reads.
    [Test]
    public async Task TemplateCellControlHoldingTableIsFollowedByParagraph()
    {
        using var template = BuildTemplate("{{ Name }}");
        AddHeader(template, OneCell(Control(OneCell("inner"))));

        using var doc = await Render(template, new());

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await AssertControlThenParagraph(header.GetFirstChild<Table>()!);
    }

    [Test]
    public async Task MarkdownFlowHeaderCellControlHoldingTableIsFollowedByParagraph()
    {
        using var styleSource = DocxTemplateBuilder.Build("replaced by the markdown");
        AddHeader(
            styleSource,
            OneCell(Control(OneCell("inner"), Paragraph("{% if Flag %}"), Paragraph("{{ Name }}"), Paragraph("{% endif %}"))));
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<CellModel>("# {{ Name }}", styleSource);

        using var doc = await Render(store, new());

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await AssertControlThenParagraph(header.GetFirstChild<Table>()!);
    }

    [Test]
    public async Task NestedControlsHoldingTableAreFollowedByOneParagraph()
    {
        var cell = new TableCell(Control(Control(OneCell("inner"))));

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Count).IsEqualTo(2);
        await Assert.That(cell.ChildElements[0]).IsTypeOf<SdtBlock>();
        await Assert.That(cell.ChildElements[1]).IsTypeOf<Paragraph>();
        await Assert.That(cell.Descendants<Paragraph>().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task NestedControlsEndingWithParagraphAreUntouched()
    {
        // The shape of an editable repeating section in a cell: a control of controls.
        var cell = new TableCell(Control(Control(OneCell("inner"), Paragraph("last"))));

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Single()).IsTypeOf<SdtBlock>();
    }

    [Test]
    public async Task CustomXmlHoldingTableIsFollowedByParagraph()
    {
        var cell = new TableCell(
            new CustomXmlBlock(new CustomXmlProperties(), OneCell("inner"))
            {
                Element = "wrapper"
            });

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Count).IsEqualTo(2);
        await Assert.That(cell.ChildElements[0]).IsTypeOf<CustomXmlBlock>();
        await Assert.That(cell.ChildElements[1]).IsTypeOf<Paragraph>();
    }

    [Test]
    public async Task CustomXmlEndingWithParagraphIsUntouched()
    {
        var cell = new TableCell(
            new CustomXmlBlock(new CustomXmlProperties(), OneCell("inner"), Paragraph("last"))
            {
                Element = "wrapper"
            });

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Single()).IsTypeOf<CustomXmlBlock>();
    }

    [Test]
    public async Task EmptyCustomXmlIsFollowedByParagraph()
    {
        // Only an empty content control is known to stand in for the paragraph of its cell.
        var cell = new TableCell(
            new CustomXmlBlock
            {
                Element = "wrapper"
            });

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Count).IsEqualTo(2);
        await Assert.That(cell.ChildElements[1]).IsTypeOf<Paragraph>();
    }

    // The editable range of a control that shares its cell ends after the control, and the
    // paragraph that closes the cell goes after that marker.
    [Test]
    public async Task ControlHoldingTableThenRangeMarkerIsFollowedByParagraph()
    {
        var cell = new TableCell(
            new PermStart
            {
                Id = 1
            },
            Control(OneCell("inner")),
            new PermEnd
            {
                Id = 1
            });

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Count).IsEqualTo(4);
        await Assert.That(cell.ChildElements[2]).IsTypeOf<PermEnd>();
        await Assert.That(cell.ChildElements[3]).IsTypeOf<Paragraph>();
    }

    [Test]
    public async Task EmptyControlAfterTableIsFollowedByParagraph()
    {
        // An empty control adds no block, so this cell still ends with its table.
        var cell = new TableCell(OneCell("inner"), Control());

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Count).IsEqualTo(3);
        await Assert.That(cell.ChildElements[2]).IsTypeOf<Paragraph>();
    }

    [Test]
    public async Task EmptyControlAfterParagraphIsUntouched()
    {
        var cell = new TableCell(Paragraph("first"), Control());

        TableCells.EnsureFinalParagraph(new Body(Wrap(cell)));

        await Assert.That(cell.ChildElements.Count).IsEqualTo(2);
        await Assert.That(cell.ChildElements[1]).IsTypeOf<SdtBlock>();
    }

    static TableCell Cell(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!.GetFirstChild<Table>()!.Descendants<TableCell>().First();

    static async Task AssertControlThenParagraph(WordprocessingDocument doc)
    {
        await AssertCellsEndWithParagraph(doc);
        await AssertControlThenParagraph(doc.MainDocumentPart!.Document!.Body!.GetFirstChild<Table>()!);
    }

    /// <summary>
    /// The first cell of <paramref name="table"/> holds a control whose content is one table, then
    /// one empty paragraph.
    /// </summary>
    static async Task AssertControlThenParagraph(Table table)
    {
        var children = table.Descendants<TableCell>().First().ChildElements;
        await Assert.That(children.Count).IsEqualTo(2);
        var control = (SdtBlock) children[0];
        await Assert.That(control.SdtContentBlock!.ChildElements.Single()).IsTypeOf<Table>();
        await Assert.That(children[1]).IsTypeOf<Paragraph>();
        await Assert.That(children[1].InnerText).IsEqualTo("");
    }

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

    static Task<WordprocessingDocument> Render(CellModel model, params OpenXmlElement[] cellBlocks)
    {
        using var template = BuildTemplate(cellBlocks);
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
    static MemoryStream BuildTemplate(params string[] cellParagraphs) =>
        BuildTemplate(cellParagraphs.Select(Paragraph).ToArray<OpenXmlElement>());

    /// <summary>
    /// A body holding a one-cell table whose cell has the given blocks, then a paragraph.
    /// </summary>
    static MemoryStream BuildTemplate(params OpenXmlElement[] cellBlocks)
    {
        var stream = DocxTemplateBuilder.Build("after");
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            doc.MainDocumentPart!.Document!.Body!.PrependChild(OneCell(cellBlocks));
        }

        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Gives <paramref name="template"/> a default header holding <paramref name="table"/>, then a
    /// paragraph.
    /// </summary>
    static void AddHeader(MemoryStream template, Table table)
    {
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new(table, new Paragraph());
            mainPart.Document!.Body!.GetFirstChild<SectionProperties>()!.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = mainPart.GetIdOfPart(headerPart)
                });
        }

        template.Position = 0;
    }

    static Table OneCell(params string[] cellParagraphs) =>
        OneCell(cellParagraphs.Select(Paragraph).ToArray<OpenXmlElement>());

    static Table OneCell(params OpenXmlElement[] cellBlocks) =>
        Wrap(new(cellBlocks));

    static Table Wrap(TableCell cell) =>
        new(
            new TableProperties(
                new TableWidth
                {
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(new GridColumn()),
            new TableRow(cell));

    static Paragraph Paragraph(string text) =>
        new(new Run(new Text(text)));

    /// <summary>
    /// A block-level content control, as a template author places one around part of a cell.
    /// </summary>
    static SdtBlock Control(params OpenXmlElement[] content) =>
        new(
            new SdtProperties(
                new SdtId
                {
                    Val = 8
                }),
            new SdtContentBlock(content));
}
