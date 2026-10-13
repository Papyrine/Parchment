// ReSharper disable PartialTypeWithSinglePart
using DocumentFormat.OpenXml.Validation;

/// <summary>
/// Word merges two <c>w:tbl</c> siblings into one table. Tables that are part of the template are
/// kept apart there by the paragraphs between them, and the render removes some of those: block
/// tag paragraphs, a branch that was not chosen, a token that produced nothing. Each test asserts
/// the sequence of blocks left behind.
/// </summary>
public partial class StaticTableSeparationTests
{
    [ParchmentBindable]
    public partial class Report
    {
        public bool Flag { get; init; }
        public IReadOnlyList<Item> Items { get; init; } = [];

        [ExcelsiorTable]
        public IReadOnlyList<Item>? Rows { get; init; }
    }

    public class Item
    {
        public required string Name { get; init; }
        public bool Active { get; init; }
    }

    static Report Flagged =>
        new()
        {
            Flag = true
        };

    static Report WithItems(params string[] names) =>
        new()
        {
            Items = names
                .Select(_ => new Item
                {
                    Name = _
                })
                .ToList()
        };

    [Test]
    public async Task TablesAroundAChosenBranchStaySeparate()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildParagraph("{% if Flag %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildTable("C"));

        using var doc = await Render(template, Flagged);

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A p: tbl:B p: tbl:C");
    }

    [Test]
    public async Task TablesAroundARemovedBranchStaySeparate()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildParagraph("{% if Flag %}"),
            BuildParagraph("Only when flagged"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildTable("C"));

        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A p: tbl:C");
    }

    [Test]
    [Arguments(true, "tbl:A p: tbl:Yes p: tbl:C")]
    [Arguments(false, "tbl:A p: tbl:No p: tbl:C")]
    public async Task TablesInEitherBranchStaySeparate(bool flag, string expected)
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildParagraph("{% if Flag %}"),
            BuildTable("Yes"),
            BuildParagraph("{% else %}"),
            BuildTable("No"),
            BuildParagraph("{% endif %}"),
            BuildTable("C"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        await Assert.That(Blocks(Body(doc))).IsEqualTo(expected);
    }

    [Test]
    public async Task RemovedBranchTableLeavesNoParagraphBehind()
    {
        // The kept table and the removed one are briefly neighbours while the else tag between
        // them is taken out. Nothing is left beside the kept table, so nothing is added.
        using var template = BuildTemplate(
            BuildParagraph("{% if Flag %}"),
            BuildTable("Yes"),
            BuildParagraph("{% else %}"),
            BuildTable("No"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("After"));

        using var doc = await Render(template, Flagged);

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:Yes p:After");
    }

    [Test]
    public async Task LoopOverAWholeTableKeepsEachCopySeparate()
    {
        using var template = BuildTemplate(
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"));

        using var doc = await Render(template, WithItems("one", "two", "three"));

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:one p: tbl:two p: tbl:three");
    }

    [Test]
    public async Task LoopedTableStaysSeparateFromTheTablesAroundTheLoop()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"),
            BuildTable("C"));

        using var doc = await Render(template, WithItems("one", "two"));

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A p: tbl:one p: tbl:two p: tbl:C");
    }

    [Test]
    public async Task TablesAroundAnEmptyLoopStaySeparate()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"),
            BuildTable("C"));

        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A p: tbl:C");
    }

    [Test]
    public async Task LoopBodyThatEndsAndStartsWithATableStaysSeparateAcrossIterations()
    {
        using var template = BuildTemplate(
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("Between"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"));

        using var doc = await Render(template, WithItems("one", "two"));

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:one p:Between tbl:one p: tbl:two p:Between tbl:two");
    }

    [Test]
    public async Task ConditionalTableInsideALoopStaysSeparateAcrossIterations()
    {
        using var template = BuildTemplate(
            BuildParagraph("{% for item in Items %}"),
            BuildParagraph("{% if item.Active %}"),
            BuildTable("Active {{ item.Name }}"),
            BuildParagraph("{% else %}"),
            BuildTable("Inactive {{ item.Name }}"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% endfor %}"));

        using var doc = await Render(
            template,
            new()
            {
                Items =
                [
                    new()
                    {
                        Name = "one",
                        Active = true
                    },
                    new()
                    {
                        Name = "two"
                    }
                ]
            });

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:Active one p: tbl:Inactive two");
    }

    [Test]
    public async Task TablesAroundATokenThatProducedNothingStaySeparate()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildParagraph("{{ Rows }}"),
            BuildTable("C"));

        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A p: tbl:C");
    }

    [Test]
    public async Task TablesAdjacentInTheTemplateAreLeftAlone()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            BuildTable("B"),
            BuildParagraph("{% if Flag %}"),
            BuildTable("C"),
            BuildTable("D"),
            BuildParagraph("{% endif %}"));

        using var doc = await Render(template, Flagged);

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A tbl:B p: tbl:C tbl:D");
    }

    [Test]
    public async Task TablesAdjacentInALoopBodyAreLeftAloneWithinAnIteration()
    {
        using var template = BuildTemplate(
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildTable("Second"),
            BuildParagraph("{% endfor %}"));

        using var doc = await Render(template, WithItems("one", "two"));

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:one tbl:Second p: tbl:two tbl:Second");
    }

    [Test]
    public async Task TableBesideParagraphsGetsNoSeparator()
    {
        using var template = BuildTemplate(
            BuildParagraph("Before"),
            BuildParagraph("{% if Flag %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% for item in Items %}"),
            BuildParagraph("{{ item.Name }}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = true,
                Items =
                [
                    new()
                    {
                        Name = "one"
                    },
                    new()
                    {
                        Name = "two"
                    }
                ]
            });

        await Assert.That(Blocks(Body(doc))).IsEqualTo("p:Before tbl:B p:one tbl:one p:two tbl:two");
    }

    [Test]
    public async Task TableEndingTheBodyGetsNoParagraphAfterIt()
    {
        // The section properties follow the last table directly, as they did before.
        using var template = BuildTemplate(
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"));

        using var doc = await Render(template, WithItems("one"));

        var body = Body(doc);
        await Assert.That(Blocks(body)).IsEqualTo("tbl:one");
        await Assert.That(body.LastChild).IsTypeOf<SectionProperties>();
        await Assert.That(body.LastChild!.PreviousSibling()).IsTypeOf<Table>();
    }

    [Test]
    public async Task RangeMarkersBetweenTheBlocksAreLookedThrough()
    {
        using var template = BuildTemplate(
            BuildTable("A"),
            new BookmarkStart
            {
                Id = "900",
                Name = "span"
            },
            BuildParagraph("{% if Flag %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            new BookmarkEnd
            {
                Id = "900"
            },
            BuildTable("C"));

        using var doc = await Render(template, Flagged);

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A bookmarkStart p: tbl:B bookmarkEnd p: tbl:C");
    }

    [Test]
    public async Task NestedTablesAroundAChosenBranchStaySeparate()
    {
        using var template = BuildTemplate(
            BuildTable(
                BuildTable("A"),
                BuildParagraph("{% if Flag %}"),
                BuildTable("B"),
                BuildParagraph("{% endif %}"),
                BuildTable("C"),
                BuildParagraph("End of cell")),
            BuildParagraph("After"));

        using var doc = await Render(template, Flagged);

        await Assert.That(Blocks(OuterCell(doc))).IsEqualTo("tbl:A p: tbl:B p: tbl:C p:End of cell");
    }

    [Test]
    public async Task NestedTableLoopKeepsEachCopySeparate()
    {
        // The cell ends with the last copy, so it is closed by the one paragraph Word requires
        // there and nothing more.
        using var template = BuildTemplate(
            BuildTable(
                BuildParagraph("{% for item in Items %}"),
                BuildTable("{{ item.Name }}"),
                BuildParagraph("{% endfor %}")),
            BuildParagraph("After"));

        using var doc = await Render(template, WithItems("one", "two"));

        await Assert.That(Blocks(OuterCell(doc))).IsEqualTo("tbl:one p: tbl:two p:");
    }

    [Test]
    public async Task NestedTableEndingACellIsClosedByOneParagraph()
    {
        using var template = BuildTemplate(
            BuildTable(
                BuildParagraph("{% if Flag %}"),
                BuildTable("B"),
                BuildParagraph("{% endif %}")),
            BuildParagraph("After"));

        using var doc = await Render(template, Flagged);

        await Assert.That(Blocks(OuterCell(doc))).IsEqualTo("tbl:B p:");
    }

    [Test]
    public async Task HeaderTablesAroundAChosenBranchStaySeparate()
    {
        using var template = BuildTemplate(BuildParagraph("Body"));
        AddHeader(
            template,
            BuildTable("A"),
            BuildParagraph("{% if Flag %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% for item in Items %}"),
            BuildTable("{{ item.Name }}"),
            BuildParagraph("{% endfor %}"),
            BuildParagraph("End of header"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = true,
                Items =
                [
                    new()
                    {
                        Name = "one"
                    },
                    new()
                    {
                        Name = "two"
                    }
                ]
            });

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await Assert.That(Blocks(header)).IsEqualTo("tbl:A p: tbl:B p: tbl:one p: tbl:two p:End of header");
    }

    [Test]
    public async Task MarkdownTemplateHeaderTablesStaySeparate()
    {
        // A markdown template's header comes from its style source and runs the docx flow.
        using var styleSource = BuildTemplate();
        AddHeader(
            styleSource,
            BuildTable("A"),
            BuildParagraph("{% if Flag %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildTable("C"),
            BuildParagraph("End of header"));

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Report>("Body", styleSource);
        using var doc = await Render(store, Flagged);

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await Assert.That(Blocks(header)).IsEqualTo("tbl:A p: tbl:B p: tbl:C p:End of header");
    }

    [Test]
    public async Task OutputValidates()
    {
        using var template = BuildTemplate(
            BuildTable(
                BuildParagraph("{% for item in Items %}"),
                BuildTable("{{ item.Name }}"),
                BuildParagraph("{% endfor %}")),
            BuildParagraph("{% if Flag %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildTable("C"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = true,
                Items =
                [
                    new()
                    {
                        Name = "one"
                    },
                    new()
                    {
                        Name = "two"
                    }
                ]
            });

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static Body Body(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!;

    static TableCell OuterCell(WordprocessingDocument doc) =>
        Body(doc).GetFirstChild<Table>()!.GetFirstChild<TableRow>()!.GetFirstChild<TableCell>()!;

    /// <summary>
    /// The children of a container as <c>name:text</c>, leaving out the properties element a body
    /// ends with or a cell starts with.
    /// </summary>
    static string Blocks(OpenXmlElement container) =>
        string.Join(
            ' ',
            container.ChildElements
                .Where(_ => _ is not (SectionProperties or TableCellProperties))
                .Select(Describe));

    static string Describe(OpenXmlElement element)
    {
        if (element is Paragraph or Table)
        {
            return $"{element.LocalName}:{element.InnerText}";
        }

        return element.LocalName;
    }

    static Task<WordprocessingDocument> Render(Stream template, Report model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Report>(template);
        return Render(store, model);
    }

    static async Task<WordprocessingDocument> Render(TemplateStore store, Report model)
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    static MemoryStream BuildTemplate(params OpenXmlElement[] blocks)
    {
        var stream = DocxTemplateBuilder.Build();
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var section = body.GetFirstChild<SectionProperties>()!;
            foreach (var block in blocks)
            {
                body.InsertBefore(block, section);
            }
        }

        stream.Position = 0;
        return stream;
    }

    static void AddHeader(MemoryStream template, params OpenXmlElement[] blocks)
    {
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(blocks);
            mainPart.Document!.Body!.GetFirstChild<SectionProperties>()!.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdHeader"
                });
        }

        template.Position = 0;
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    static Table BuildTable(string text) =>
        BuildTable(BuildParagraph(text));

    static Table BuildTable(params OpenXmlElement[] cellBlocks) =>
        new(
            new TableProperties(
                new TableWidth
                {
                    Width = "0",
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(new GridColumn()),
            new TableRow(new TableCell(cellBlocks)));
}
