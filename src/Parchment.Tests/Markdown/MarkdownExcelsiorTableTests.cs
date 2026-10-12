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

    [ParchmentBindable]
    public partial class ShadowedQuoteModel
    {
        public required QuoteModel Order { get; init; }
        public required List<OrderSummary> Orders { get; init; }
    }

    public class OrderSummary
    {
        public required int Lines { get; init; }
    }

    // `order.Lines` spells the root's `Order.Lines`, but inside the loop it is the item's member.
    // Taken for the table it rendered the root's rows once per item, and the second token was
    // refused at registration for sharing its line.
    [Test]
    public async Task LoopVariableNamedAfterRootTableReadsTheItem()
    {
        var markdown =
            """
            {% for order in Orders %}
            {{ order.Lines }}

            Next: {{ order.Lines | plus: 1 }}

            {% endfor %}
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ShadowedQuoteModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ShadowedQuoteModel
            {
                Order = Quote(),
                Orders =
                [
                    new()
                    {
                        Lines = 3
                    },
                    new()
                    {
                        Lines = 5
                    }
                ]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0);

        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        await Assert.That(paragraphs).IsEquivalentTo(["3", "Next: 4", "5", "Next: 6"]);
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

    static async Task<Body> RenderShadowed(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ShadowedQuoteModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ShadowedQuoteModel
            {
                Order = Quote(),
                Orders =
                [
                    new()
                    {
                        Lines = 3
                    }
                ]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body) doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    // Whitespace control and a line break inside the braces are spellings of the same token. They
    // used to miss the rewrite and print the row type's name once per row.
    [Test]
    [Arguments("{{- Lines -}}")]
    [Arguments("{{- Lines }}")]
    [Arguments("{{ Lines -}}")]
    [Arguments("{{-Lines-}}")]
    [Arguments("{{\n  Lines\n}}")]
    [Arguments("  {{-\r\n  Lines\r\n-}}  ")]
    [Arguments("# Quote\n\n{{ Lines -}}")]
    [Arguments("{{- Lines }}\n\nThank you.")]
    [Arguments("# Quote\r\n\r\n{{\r\n  Lines\r\n}}\r\n\r\nThank you.")]
    public async Task EverySpellingRendersAsAWordTable(string markdown)
    {
        using var stream = await Render(markdown);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.Descendants<Table>().Single().InnerText).Contains("Site survey");

        var outside = string.Concat(body.Elements<Paragraph>().Select(_ => _.InnerText));
        await Assert.That(outside.Replace("Quote", "").Replace("Thank you.", "")).IsEqualTo("");
    }

    // `Order['Lines']` is `Order.Lines` to Fluid, so it is the same table.
    [Test]
    [Arguments("{{ Order['Lines'] }}")]
    [Arguments("{{ Order[\"Lines\"] }}")]
    [Arguments("{{- Order['Lines'] -}}")]
    [Arguments("{{\n  Order['Lines']\n}}")]
    public async Task BracketAccessRendersAsAWordTable(string markdown)
    {
        var body = await RenderShadowed(markdown);

        await Assert.That(body.Descendants<Table>().Single().InnerText).Contains("Site survey");
        await Assert.That(string.Concat(body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("");
    }

    // The marker inherits the token's dashes. Between tags that is what keeps it a block of its
    // own; against text it joins the text's paragraph, which the swap refuses at render.
    [Test]
    public async Task TableTokenKeepsItsWhitespaceControl()
    {
        using var stream = await Render("Before\n\n{% if Reference %}\n\n   {{- Lines -}}   \n\n{% endif %}\n\nAfter");
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.Descendants<Table>().Count()).IsEqualTo(1);
        await Assert.That(body.Elements<Paragraph>().Select(_ => _.InnerText).Where(_ => _.Length > 0))
            .IsEquivalentTo(["Before", "After"]);

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<QuoteModel>("Before\n\n{{- Lines }}", styleSource);
        using var target = new MemoryStream();
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(() => store.Render(Quote(), target));
        await Assert.That(exception!.Message).Contains("has to be alone in its block");
    }

    // The loop variable still wins in the other spellings, alone on its line or not.
    [Test]
    [Arguments("{{- order.Lines -}}")]
    [Arguments("{{ order['Lines'] }}")]
    [Arguments("{{\n  order.Lines\n}}")]
    [Arguments("{{- order.Lines | plus: 0 -}}")]
    [Arguments("{{ order.Lines }}{{- '' -}}")]
    public async Task LoopVariableReadsTheItemInEverySpelling(string token)
    {
        var body = await RenderShadowed($"{{% for order in Orders %}}\n\n{token}\n\n{{% endfor %}}");

        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        await Assert.That(body.InnerText).IsEqualTo("3");
    }

    [Test]
    [Arguments("Lines: {{- Order.Lines -}}")]
    [Arguments("{{ Order['Lines'] }} follow")]
    [Arguments("{{ Order['Lines'] }}{{ Order.Lines }}")]
    [Arguments("Lines: {{\n  Order.Lines\n}}")]
    [Arguments("{{\n  Order.Lines\n}} follow")]
    [Arguments("Count: {{ Order['Lines'] | size }}")]
    public async Task TokenSharingItsLineIsRejectedInEverySpelling(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();

        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => store.RegisterMarkdownTemplate<ShadowedQuoteModel>(markdown, styleSource));

        await Assert.That(exception.Message)
            .Contains("'{{ Order.Lines }}' renders an [ExcelsiorTable] as a Word table, which replaces the whole block it sits in, so it has to be alone on its line.");
    }

    [Test]
    [Arguments("{{- Order.Lines | reverse -}}")]
    [Arguments("{{ Order['Lines'] | size }}")]
    [Arguments("{{\n  Order.Lines\n  | reverse\n}}")]
    [Arguments("{{ Order['Lines'][0] }}")]
    public async Task FilteredTokenIsRejectedInEverySpelling(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();

        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => store.RegisterMarkdownTemplate<ShadowedQuoteModel>(markdown, styleSource));

        await Assert.That(exception.Message)
            .Contains($"'{markdown}' renders the [ExcelsiorTable] 'Order.Lines' as a Word table");
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
