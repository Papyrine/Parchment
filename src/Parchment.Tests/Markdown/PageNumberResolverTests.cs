
// ReSharper disable PartialTypeWithSinglePart

public partial class PageNumberResolverTests
{
    [ParchmentBindable]
    public partial class ReportModel
    {
        public required string Heading { get; init; }
    }

    const string markdown =
        """
        Contents {.TOCHeading}

        [TOC]{levels=1}

        # {{ Heading }}

        Body text.

        # Risks {#risks}

        See the summary on page [](#risks).
        """;

    // Whatever paginates the document reports where each bookmark landed; everything else — the
    // entries, the styles, the links — is Parchment's to build. A limit stands in for an engine that
    // could not place every bookmark.
    class StubResolver(int limit = int.MaxValue) : IPageNumberResolver
    {
        public List<string> Bookmarks { get; } = [];

        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel)
        {
            using var doc = WordprocessingDocument.Open(docx, false);
            var names = doc.MainDocumentPart!.Document!.Body!
                .Descendants<BookmarkStart>()
                .Select(_ => _.Name?.Value)
                .Where(_ => _ != null)
                .ToList();
            Bookmarks.AddRange(names!);

            IReadOnlyDictionary<string, int> pages = names
                .Take(limit)
                .Select((name, index) => (name: name!, page: index + 2))
                .ToDictionary(_ => _.name, _ => _.page, StringComparer.Ordinal);
            return Task.FromResult(pages);
        }
    }

    static async Task<MemoryStream> RenderStream(IPageNumberResolver? resolver, string template = markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = resolver
        };
        store.RegisterMarkdownTemplate<ReportModel>(template, styleSource);

        var stream = new MemoryStream();
        await store.Render(
            new ReportModel
            {
                Heading = "Delivery"
            },
            stream);
        stream.Position = 0;
        return stream;
    }

    static async Task<Body> Render(IPageNumberResolver? resolver, string template = markdown)
    {
        using var stream = await RenderStream(resolver, template);
        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    [Test]
    public async Task EntriesAreBuiltFromTheHeadings()
    {
        var body = await Render(new StubResolver());
        var entries = body.Descendants<Hyperlink>().ToList();

        await Assert.That(entries.Count).IsEqualTo(2);
        await Assert.That(Text(entries[0])).IsEqualTo("Delivery");
        await Assert.That(Text(entries[1])).IsEqualTo("Risks");
    }

    [Test]
    public async Task EntriesTakeTheMatchingTocStyle()
    {
        var body = await Render(new StubResolver());
        var entry = body.Elements<Paragraph>().First(_ => _.Elements<Hyperlink>().Any());

        await Assert.That(entry.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("TOC1");
        var tab = entry.ParagraphProperties.GetFirstChild<Tabs>()!.GetFirstChild<TabStop>()!;
        await Assert.That(tab.Val?.Value).IsEqualTo(TabStopValues.Right);
        await Assert.That(tab.Leader?.Value).IsEqualTo(TabStopLeaderCharValues.Dot);
    }

    // The field's begin/instruction/separate markers move onto the first entry, and a w:pPr that is
    // no longer the paragraph's first child is one Word throws away whole: that entry renders as
    // body text with no dot leader while every entry below it looks right.
    [Test]
    public async Task FieldMarkersLandAfterTheFirstEntrysProperties()
    {
        using var stream = await RenderStream(new StubResolver());
        using var doc = WordprocessingDocument.Open(stream, false);
        var entry = doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .First(_ => _.Elements<Hyperlink>().Any());

        await Assert.That(entry.Descendants<FieldChar>().Any()).IsTrue();
        await Assert.That(entry.FirstChild).IsTypeOf<ParagraphProperties>();

        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    // A heading the author never gave an id still has to be linkable, so it gets one named the way
    // Word names its own.
    [Test]
    public async Task HeadingsWithoutAnIdAreBookmarked()
    {
        var resolver = new StubResolver();
        await Render(resolver);

        await Assert.That(resolver.Bookmarks).Contains(_ => _.StartsWith("_Toc"));
        await Assert.That(resolver.Bookmarks).Contains("risks");
    }

    [Test]
    public async Task ResolvedPageNumbersReplaceThePlaceholders()
    {
        var body = await Render(new StubResolver());
        var entries = body.Descendants<Hyperlink>().ToList();

        // The stub numbers bookmarks in document order from page 2.
        await Assert.That(PageNumber(entries[0])).IsEqualTo("2");
        await Assert.That(PageNumber(entries[1])).IsEqualTo("3");
    }

    [Test]
    public async Task PageReferenceCarriesTheResolvedNumber()
    {
        var body = await Render(new StubResolver());
        var field = body.Descendants<FieldCode>().Single(_ => _.Text.Contains("PAGEREF"));
        var paragraph = field.Ancestors<Paragraph>().First();
        // The result run is what Word displays: the text after the field's separate marker.
        var displayed = paragraph.Elements<Run>()
            .SkipWhile(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Separate)
            .Skip(1)
            .First();

        // "risks" is the second bookmark, so the stub puts it on page 3.
        await Assert.That(displayed.GetFirstChild<Text>()!.Text).IsEqualTo("3");
    }

    // The whole point of resolving: Word has nothing left to recalculate, so it opens without
    // asking the reader for permission to do it.
    [Test]
    public async Task ResolvedFieldsAreNoLongerDirty()
    {
        var body = await Render(new StubResolver());

        await Assert.That(body.Descendants<FieldChar>().Any(_ => _.Dirty?.Value == true)).IsFalse();
    }

    [Test]
    public async Task WithoutAResolverWordIsStillAskedToBuildIt()
    {
        var body = await Render(null);

        await Assert.That(body.Descendants<Hyperlink>().Any(_ => _.Anchor?.Value?.StartsWith("_Toc") == true)).IsFalse();
        await Assert.That(body.Descendants<FieldChar>().Any(_ => _.Dirty?.Value == true)).IsTrue();
    }

    const string headings =
        """


        # First

        # Second
        """;

    const string pipeCell =
        """
        | Contents |
        | --- |
        | [TOC] |
        """;

    const string pipeHeaderCell =
        """
        | [TOC] |
        | --- |
        | Contents |
        """;

    const string gridCell =
        """
        +----------+
        | Contents |
        +==========+
        | [TOC]    |
        +----------+
        """;

    const string cellInAQuote =
        """
        > | Contents |
        > | --- |
        > | [TOC] |
        """;

    const string cellInAListItem =
        """
        - Item

          | Contents |
          | --- |
          | [TOC] |
        """;

    const string cellInACell =
        """
        +----------------+
        | Outer          |
        +================+
        | | Contents |   |
        | | --- |        |
        | | [TOC] |      |
        +----------------+
        """;

    // A table cell is the one place the renderer writes a [TOC] below the body: a quote or a list
    // item flattens to body paragraphs, but a table inside either still has cells, and so does a
    // table inside another table's cell. The entries are built where the field is, so that is where
    // their numbers have to be written.
    [Test]
    [Arguments(pipeCell, 1)]
    [Arguments(pipeHeaderCell, 1)]
    [Arguments(gridCell, 1)]
    [Arguments(cellInAQuote, 1)]
    [Arguments(cellInAListItem, 1)]
    [Arguments(cellInACell, 2)]
    public async Task EntriesInATableCellTakeTheResolvedPageNumbers(string table, int depth)
    {
        using var stream = await RenderStream(new StubResolver(), table + headings);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var entries = body.Descendants<Hyperlink>().ToList();

        await Assert.That(entries.Count).IsEqualTo(2);
        await Assert.That(entries[0].Ancestors<TableCell>().Count()).IsEqualTo(depth);
        await Assert.That(entries[1].Ancestors<TableCell>().Count()).IsEqualTo(depth);

        // The stub numbers bookmarks in document order from page 2.
        await Assert.That(PageNumber(entries[0])).IsEqualTo("2");
        await Assert.That(PageNumber(entries[1])).IsEqualTo("3");
        await Assert.That(body.Descendants<FieldChar>().Any(_ => _.Dirty?.Value == true)).IsFalse();

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    // Each field's entries are found from the field itself, so one in the body and one in a cell do
    // not answer for each other.
    [Test]
    public async Task ATableOfContentsInTheBodyAndOneInACellAreBothResolved()
    {
        var body = await Render(new StubResolver(), "[TOC]\n\n" + pipeCell + headings);
        var entries = body.Descendants<Hyperlink>().ToList();

        await Assert.That(entries.Select(_ => _.Ancestors<TableCell>().Count())).IsEquivalentTo([0, 0, 1, 1]);
        await Assert.That(entries.Select(PageNumber)).IsEquivalentTo(["2", "3", "2", "3"]);
        await Assert.That(body.Descendants<FieldChar>().Any(_ => _.Dirty?.Value == true)).IsFalse();
    }

    // An entry the resolver could not place keeps its placeholder, and the field stays marked so
    // Word still offers the reader a number — which it cannot do for a field that was cleared
    // because its entries were never looked at.
    [Test]
    [Arguments("[TOC]")]
    [Arguments(pipeCell)]
    public async Task AnEntryTheResolverCouldNotPlaceLeavesTheFieldDirty(string contents)
    {
        var body = await Render(new StubResolver(limit: 1), contents + headings);
        var entries = body.Descendants<Hyperlink>().ToList();
        var begin = body.Descendants<FieldChar>().Single(_ => _.FieldCharType?.Value == FieldCharValues.Begin);

        await Assert.That(PageNumber(entries[0])).IsEqualTo("2");
        await Assert.That(PageNumber(entries[1])).IsEqualTo("1");
        await Assert.That(begin.Dirty?.Value).IsTrue();
    }

    // Left for Word, a field in a cell is the field the body gets: the same runs, still dirty, and no
    // entries built for it.
    [Test]
    public async Task WithoutAResolverAFieldInACellIsTheOneTheBodyGets()
    {
        var inBody = await Render(null, "[TOC]" + headings);
        var inCell = await Render(null, pipeCell + headings);

        var bodyField = inBody.Elements<Paragraph>().First();
        var cellField = inCell.Descendants<TableCell>().Last().Elements<Paragraph>().Single();

        await Assert.That(cellField.OuterXml).IsEqualTo(bodyField.OuterXml);
        await Assert.That(cellField.Descendants<FieldChar>().First().Dirty?.Value).IsTrue();
        await Assert.That(inCell.Descendants<Hyperlink>().Any()).IsFalse();
    }

    static string Text(Hyperlink hyperlink) =>
        hyperlink.Elements<Run>().First().GetFirstChild<Text>()!.Text;

    static string PageNumber(Hyperlink hyperlink) =>
        hyperlink.Elements<Run>().Last().GetFirstChild<Text>()!.Text;
}
