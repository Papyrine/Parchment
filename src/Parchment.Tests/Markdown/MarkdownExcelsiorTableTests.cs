// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// <c>[ExcelsiorTable]</c> in the markdown flow: the token is replaced before the source is parsed
/// and the marker's paragraph becomes the table afterwards, so a markdown template gets the same
/// Word table the docx flow builds.
/// </summary>
public partial class MarkdownExcelsiorTableTests
{
    [ParchmentBindable]
    public partial class QuoteModel
    {
        public required string Reference { get; init; }

        [ExcelsiorTable]
        public required IReadOnlyList<QuoteLine> Lines { get; init; }
    }

    #region ExcelsiorTableConfigure

    [ParchmentBindable]
    public partial class ConfiguredQuoteModel
    {
        [ExcelsiorTable(Configure = nameof(Configure))]
        public required IReadOnlyList<QuoteLine> Lines { get; init; }

        // The escape hatch for what the attribute cannot say: it carries constants, and per-column
        // configuration is code. Called with the builder after the attribute settings are applied.
        static void Configure(WordTableBuilder<QuoteLine> builder) =>
            builder.Column(_ => _.Description, _ => _.Heading = "Deliverable");
    }

    #endregion

    public class QuoteLine
    {
        [Column(Heading = "Item", Order = 1)]
        public required string Description { get; init; }

        [Column(Heading = "Qty", Order = 2)]
        public required int Quantity { get; init; }

        [Column(Order = 3, Format = "C0")]
        public required decimal UnitPrice { get; init; }
    }

    static QuoteModel Quote() =>
        new()
        {
            Reference = "Q-1024",
            Lines =
            [
                new() { Description = "Site survey", Quantity = 1, UnitPrice = 1200m },
                new() { Description = "Cabling", Quantity = 40, UnitPrice = 35m }
            ]
        };

    // Configure is resolved by the source generator into a direct call, so this passing at all
    // proves the pipeline: attribute name -> generated delegate -> bridge -> builder.
    [Test]
    public async Task ConfigureShapesTheBuilder()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ConfiguredQuoteModel>("{{ Lines }}", styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ConfiguredQuoteModel
            {
                Lines = Quote().Lines
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var headings = doc.MainDocumentPart!.Document!.Body!
            .Descendants<Table>()
            .Single()
            .Elements<TableRow>()
            .First()
            .Elements<TableCell>()
            .Select(_ => _.InnerText)
            .ToList();

        await Assert.That(string.Join(", ", headings)).IsEqualTo("Deliverable, Qty, Unit Price");
    }

    static async Task<MemoryStream> Render(string markdown, QuoteModel? model = null)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<QuoteModel>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(model ?? Quote(), stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task RendersAsAWordTable()
    {
        var markdown =
            """
            # Quote {{ Reference }}

            {{ Lines }}

            Thank you.
            """;

        using var stream = await Render(markdown);
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class StyledQuoteModel
    {
        [ExcelsiorTable(TableStyle = "LinedColumns")]
        public required IReadOnlyList<QuoteLine> Lines { get; init; }
    }

    // The table's own look comes from the host document's styles, the same way the cell text's
    // does — without this the table falls back to Excelsior's TableGrid and a branded template
    // gets its fonts but not its borders.
    [Test]
    public async Task TableStyleReachesTheTable()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<StyledQuoteModel>("{{ Lines }}", styleSource);

        using var stream = new MemoryStream();
        await store.Render(new StyledQuoteModel { Lines = Quote().Lines }, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        var style = table.GetFirstChild<TableProperties>()!.GetFirstChild<TableStyle>()!;

        await Assert.That(style.Val?.Value).IsEqualTo("LinedColumns");
    }

    // The headings come from [Column], not from the markdown, which is the whole point of moving a
    // table onto the model: the template says where, the model says what.
    [Test]
    public async Task TakesItsColumnsFromTheModel()
    {
        using var stream = await Render("{{ Lines }}");
        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();

        await Assert.That(table.InnerText).Contains("Item");
        await Assert.That(table.InnerText).Contains("Qty");
        await Assert.That(table.InnerText).Contains("Site survey");
    }

    // A markdown pipe table and an Excelsior table have to coexist: the rewrite is keyed on the
    // token, so a hand-written table beside one is untouched.
    [Test]
    public async Task LeavesMarkdownTablesAlone()
    {
        var markdown =
            """
            | Ref |
            | --- |
            | {{ Reference }} |

            {{ Lines }}
            """;

        using var stream = await Render(markdown);
        using var doc = WordprocessingDocument.Open(stream, false);
        var tables = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().ToList();

        await Assert.That(tables.Count).IsEqualTo(2);
        await Assert.That(tables[0].InnerText).Contains("Q-1024");
        await Assert.That(tables[1].InnerText).Contains("Site survey");
    }

    // Markdig attaches a preceding {.Class} line to the block that follows it, so the marker still
    // parses into a paragraph of its own and the swap finds it.
    [Test]
    public async Task WorksUnderAGenericAttributeLine()
    {
        var markdown =
            """
            {.Caption}
            {{ Lines }}
            """;

        using var stream = await Render(markdown);
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Count()).IsEqualTo(1);
    }

    // Sharing a line means the table would have to discard the text beside it, so it is refused at
    // registration rather than rendering the collection's ToString into the document.
    [Test]
    public async Task TokenSharingItsLineIsRejected()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();

        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => store.RegisterMarkdownTemplate<QuoteModel>("Lines: {{ Lines }}", styleSource));

        await Verify(exception.Message)
            .Snapshot("Parchment registration failed for template 'QuoteModel': '{{ Lines }}' renders an [ExcelsiorTable] as a Word table, which replaces the whole block it sits in, so it has to be alone on its line. It currently shares a line with other content: Lines: {{ Lines }}");
    }

    // A filtered expression is not a path the Excelsior getter can walk, so it is refused for the
    // same reason PARCH008 refuses it in the docx flow.
    [Test]
    public async Task FilteredTokenIsRejected()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();

        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => store.RegisterMarkdownTemplate<QuoteModel>("{{ Lines | reverse }}", styleSource));

        await Verify(exception.Message)
            .Snapshot("Parchment registration failed for template 'QuoteModel': '{{ Lines | reverse }}' renders the [ExcelsiorTable] 'Lines' as a Word table, which is built from the model rather than from Fluid's output, so the token has to be a plain member access. Filters and expressions cannot be applied to it.");
    }

    // A null collection leaves no marker text behind: the paragraph goes, and no table takes its
    // place.
    [Test]
    public async Task NullCollectionRendersNothing()
    {
        var markdown =
            """
            # Quote

            {{ Lines }}
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<NullableQuote>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(new NullableQuote { Lines = null }, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        await Assert.That(body.InnerText).DoesNotContain("parchment-table");
    }

    [ParchmentBindable]
    public partial class NullableQuote
    {
        [ExcelsiorTable]
        public IReadOnlyList<QuoteLine>? Lines { get; init; }
    }
}
