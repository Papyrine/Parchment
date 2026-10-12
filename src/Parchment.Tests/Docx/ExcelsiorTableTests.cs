// ReSharper disable PartialTypeWithSinglePart
public partial class ExcelsiorTableTests
{
    static string ScenarioPath(string scenarioName) =>
        Path.GetFullPath(Path.Combine(
            ProjectFiles.ProjectDirectory,
            "Scenarios",
            scenarioName));

    #region ExcelsiorTableModel

    [ParchmentBindable]
    public partial class Quote
    {
        public required string Reference;

        [ExcelsiorTable]
        public required IReadOnlyList<QuoteLine> Lines;
    }

    public class QuoteLine
    {
        [Column(Heading = "Item", Order = 1)]
        public required string Description;

        [Column(Heading = "Qty", Order = 2)]
        public required int Quantity;

        [Column(Order = 3, Format = "C0")]
        public required decimal UnitPrice;
    }

    #endregion

    [ParchmentBindable]
    public partial class Order
    {
        public required string Number { get; init; }
        public required Buyer Buyer { get; init; }
    }

    public class Buyer
    {
        public required string Name { get; init; }

        [ExcelsiorTable]
        public required IReadOnlyList<Address> Addresses { get; init; }
    }

    public class Address
    {
        [Column(Order = 1)]
        public required string Street { get; init; }

        [Column(Order = 2)]
        public required string City { get; init; }
    }

    [ParchmentBindable]
    public partial class QuoteWithFieldLines
    {
        public required string Reference { get; init; }

        [ExcelsiorTable]
        public IReadOnlyList<QuoteLine> Lines = [];
    }

    [Test]
    public async Task FieldMarkedExcelsiorTableRendersAsTable()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Quote {{ Reference }}

            {{ Lines }}

            End.
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<QuoteWithFieldLines>(template);

        var model = new QuoteWithFieldLines
        {
            Reference = "Q-FIELD-0001",
            Lines =
            [
                new()
                {
                    Description = "Strategy workshop",
                    Quantity = 2,
                    UnitPrice = 4500m
                },
                new()
                {
                    Description = "Implementation support",
                    Quantity = 8,
                    UnitPrice = 1750m
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    #region ExcelsiorTableParagraphStyles

    [ParchmentBindable]
    public partial class StyledQuote
    {
        [ExcelsiorTable(HeadingParagraphStyle = "TBLHeading", BodyParagraphStyle = "TBLText")]
        public required IReadOnlyList<QuoteLine> Lines;
    }

    #endregion

    [Test]
    public async Task ExcelsiorTableAppliesNamedParagraphStyles()
    {
        using var template = DocxTemplateBuilder.Build("{{ Lines }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<StyledQuote>(template);

        var model = new StyledQuote
        {
            Lines =
            [
                new()
                {
                    Description = "Strategy workshop",
                    Quantity = 2,
                    UnitPrice = 4500m
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();

        var headerStyle = rows[0].GetFirstChild<TableCell>()!.GetFirstChild<Paragraph>()!
            .ParagraphProperties!.ParagraphStyleId!.Val!.Value;
        await Assert.That(headerStyle).IsEqualTo("TBLHeading");

        var bodyStyle = rows[1].GetFirstChild<TableCell>()!.GetFirstChild<Paragraph>()!
            .ParagraphProperties!.ParagraphStyleId!.Val!.Value;
        await Assert.That(bodyStyle).IsEqualTo("TBLText");
    }

    #region ExcelsiorTableViaOpenXmlToken

    [ParchmentBindable]
    public partial class GroupedReport
    {
        public required IReadOnlyList<QuoteLine> Lines;

        // [ExcelsiorTable] can't render a loop-scoped or fully-styled table, so build it directly
        // with Excelsior's WordTableBuilder and return it from an OpenXmlToken. The token's
        // {{ LinesTable }} substitution must sit alone in its paragraph (structural replacement).
        public TokenValue LinesTable =>
            new OpenXmlToken(context =>
            [
                new WordTableBuilder<QuoteLine>(Lines)
                    .HeadingParagraphStyle("TBLHeading")
                    .BodyParagraphStyle("TBLText")
                    .Build(context.MainPart)
            ]);
    }

    #endregion

    [Test]
    public async Task OpenXmlTokenWordTableBuilderBridge()
    {
        using var template = DocxTemplateBuilder.Build("{{ LinesTable }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<GroupedReport>(template);

        var model = new GroupedReport
        {
            Lines =
            [
                new()
                {
                    Description = "Strategy workshop",
                    Quantity = 2,
                    UnitPrice = 4500m
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        var bodyStyle = table.Elements<TableRow>().Skip(1).First().GetFirstChild<TableCell>()!
            .GetFirstChild<Paragraph>()!.ParagraphProperties!.ParagraphStyleId!.Val!.Value;
        await Assert.That(bodyStyle).IsEqualTo("TBLText");
    }

    [Test]
    public async Task NestedPathRendersAsExcelsiorTable()
    {
        // The [ExcelsiorTable] is on Buyer.Addresses, not on the root Order. The substitution
        // {{ Buyer.Addresses }} must walk the nested path at registration time and the runner
        // must dispatch on the dotted-path lookup.
        using var template = DocxTemplateBuilder.Build(
            """
            Order {{ Number }}

            Buyer: {{ Buyer.Name }}

            {{ Buyer.Addresses }}

            End.
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Order>(template);

        var model = new Order
        {
            Number = "ORD-2026-0001",
            Buyer = new()
            {
                Name = "Acme Corp",
                Addresses =
                [
                    new()
                    {
                        Street = "1 Pine St",
                        City = "Portland"
                    },
                    new()
                    {
                        Street = "42 Oak Ave",
                        City = "Seattle"
                    }
                ]
            }
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class Catalogue
    {
        [ExcelsiorTable]
        public required IReadOnlyList<QuoteLine> Lines;

        public required IReadOnlyList<Section> Sections;

        public string Notes = "";
    }

    public class Section
    {
        public required string Table;
    }

    const string markdownTable =
        """
        | A | B |
        |---|---|
        | 1 | 2 |
        """;

    static Catalogue BuildCatalogue() =>
        new()
        {
            Lines =
            [
                new()
                {
                    Description = "Strategy workshop",
                    Quantity = 2,
                    UnitPrice = 4500m
                }
            ],
            Sections =
            [
                new()
                {
                    Table = markdownTable
                },
                new()
                {
                    Table = markdownTable
                },
                new()
                {
                    Table = markdownTable
                }
            ],
            Notes = "Notes\n\n" + markdownTable
        };

    static async Task<List<string>> RenderBlockNames(Stream template)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Catalogue>(template);

        using var stream = new MemoryStream();
        await store.Render(BuildCatalogue(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.ChildElements
            .Where(_ => _ is not SectionProperties)
            .Select(_ => _.LocalName)
            .ToList();
    }

    static MemoryStream BuildTemplateWithBlocks(params OpenXmlElement[] blocks)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body(blocks);
            body.Append(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));
            mainPart.Document = new(body);
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    static Table BuildStaticTable() =>
        new(
            new TableRow(
                new TableCell(
                    BuildParagraph("Static"))));

    [Test]
    public async Task ConsecutiveTableTokensStaySeparateTables()
    {
        // Word merges two w:tbl siblings into one table, so the two produced tables need a
        // paragraph between them.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Lines }}

            {{ Lines }}
            """);

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("tbl p tbl");
    }

    [Test]
    public async Task TableTokenBetweenTemplateTablesStaysSeparate()
    {
        using var template = BuildTemplateWithBlocks(
            BuildStaticTable(),
            BuildParagraph("{{ Lines }}"),
            BuildStaticTable());

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("tbl p tbl p tbl");
    }

    [Test]
    public async Task TableTokenBetweenParagraphsGetsNoSeparator()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Before

            {{ Lines }}

            After
            """);

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("p tbl p");
    }

    [Test]
    public async Task TablesProducedByLoopIterationsStaySeparate()
    {
        // Each iteration's body is nothing but a produced table, and the loop sits between two
        // template tables once its tag paragraphs are removed.
        using var template = BuildTemplateWithBlocks(
            BuildStaticTable(),
            BuildParagraph("{% for section in Sections %}"),
            BuildParagraph("{{ section.Table | markdown }}"),
            BuildParagraph("{% endfor %}"),
            BuildStaticTable());

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("tbl p tbl p tbl p tbl p tbl");
    }

    [Test]
    public async Task TableProducedInsideIfStaysSeparateFromFollowingTable()
    {
        using var template = BuildTemplateWithBlocks(
            BuildParagraph("{% if Notes %}"),
            BuildParagraph("{{ Lines }}"),
            BuildParagraph("{% endif %}"),
            BuildStaticTable());

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("tbl p tbl");
    }

    [Test]
    public async Task MarkdownEndingInTableStaysSeparateFromFollowingTable()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Notes | markdown }}

            {{ Lines }}
            """);

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("p tbl p tbl");
    }

    [Test]
    public async Task NonSoloMarkdownEndingInTableKeepsOneParagraphBeforeFollowingTable()
    {
        // The split path keeps the host as the before half and always emits the after half, so
        // that paragraph already separates the tables and no second one is added.
        using var template = DocxTemplateBuilder.Build(
            """
            Intro {{ Notes | markdown }}

            {{ Lines }}
            """);

        var blocks = await RenderBlockNames(template);

        await Assert.That(string.Join(' ', blocks)).IsEqualTo("p p tbl p tbl");
    }

    [Test]
    public async Task MixedInlineContentInSameParagraphIsRejected()
    {
        // "Prefix {{ Lines }}" — the token doesn't cover the whole paragraph. Structural
        // replacement would drop "Prefix ", so registration must fail up-front.
        using var template = DocxTemplateBuilder.Build(
            """
            Quote {{ Reference }}

            Prefix {{ Lines }}

            End.
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(() => store.RegisterDocxTemplate<Quote>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("must sit alone in its own paragraph");
    }

    [Test]
    public async Task FilterOnExcelsiorTokenIsRejected()
    {
        // {{ Lines | reverse }} — the Excelsior renderer walks the model object directly
        // and bypasses Fluid, so a filter would be silently dropped. Registration must fail.
        using var template = DocxTemplateBuilder.Build(
            """
            Quote {{ Reference }}

            {{ Lines | reverse }}

            End.
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(() => store.RegisterDocxTemplate<Quote>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("plain member-access");
    }

    [Test]
    public async Task Render()
    {
        #region ExcelsiorTableUsage

        var templatePath = Path.Combine(ScenarioPath("excelsior-table"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Quote>(templatePath);

        var model = new Quote
        {
            Reference = "Q-2026-0042",
            Lines =
            [
                new()
                {
                    Description = "Strategy workshop",
                    Quantity = 2,
                    UnitPrice = 4500m
                },
                new()
                {
                    Description = "Implementation support",
                    Quantity = 8,
                    UnitPrice = 1750m
                },
                new()
                {
                    Description = "Documentation review",
                    Quantity = 1,
                    UnitPrice = 950m
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);

        #endregion

        // Land the Verify artifacts next to input.docx so the scenario directory is a
        // self-contained folder of inputs and outputs that the readme can link to directly.
        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("excelsior-table"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }
}
