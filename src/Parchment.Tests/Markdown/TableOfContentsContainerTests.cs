
// ReSharper disable PartialTypeWithSinglePart

// A table cell is the one container the markdown flow writes below the body, and a table of
// contents meets it from both sides: as where a heading it has to list can be, and as where the
// field itself can be, with the cell rather than the page as the room its entries have.
public partial class TableOfContentsContainerTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

    // DocxTemplateBuilder's page is 6500 wide with 500 on every edge.
    const int textWidth = 5500;

    // What TableRenderer puts either side of a cell's text.
    const int cellMargins = 216;

    // Numbers the bookmarks in document order from page 2.
    class StubResolver : IPageNumberResolver
    {
        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel)
        {
            using var doc = WordprocessingDocument.Open(docx, false);
            IReadOnlyDictionary<string, int> pages = doc.MainDocumentPart!.Document!.Body!
                .Descendants<BookmarkStart>()
                .Select((bookmark, index) => (name: bookmark.Name!.Value!, page: index + 2))
                .ToDictionary(_ => _.name, _ => _.page, StringComparer.Ordinal);
            return Task.FromResult(pages);
        }
    }

    const string headingInACell =
        """
        +-------------+
        | # Inside    |
        |             |
        | Cell text.  |
        +-------------+
        """;

    const string headingInACellInACell =
        """
        +--------------------+
        | Outer              |
        +====================+
        | +-------------+    |
        | | # Inside    |    |
        | |             |    |
        | | Cell text.  |    |
        | +-------------+    |
        +--------------------+
        """;

    [Test]
    [Arguments(headingInACell, 1)]
    [Arguments(headingInACellInACell, 2)]
    public async Task AHeadingInATableCellIsListedWhereItFalls(string table, int depth)
    {
        using var stream = await RenderStream(new StubResolver(), $"[TOC]\n\n# Before\n\n{table}\n\n# After\n");
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var entries = body.Descendants<Hyperlink>().ToList();

        await Assert.That(entries.Select(Text)).IsEquivalentTo(["Before", "Inside", "After"]);
        await Assert.That(entries.Select(PageNumber)).IsEquivalentTo(["2", "3", "4"]);

        // The entry links to the heading in the cell, which was bookmarked there.
        var target = body.Descendants<BookmarkStart>().Single(_ => _.Name?.Value == entries[1].Anchor?.Value);
        await Assert.That(target.Ancestors<TableCell>().Count()).IsEqualTo(depth);
        await Assert.That(body.Descendants<FieldChar>().Any(_ => _.Dirty?.Value == true)).IsFalse();

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    [Test]
    public async Task AHeadingInACellBelowTheRequestedLevelsIsLeftOut()
    {
        var body = await Render(
            new StubResolver(),
            """
            [TOC]{levels=1}

            # Before

            +-------------+
            | ## Inside   |
            +-------------+
            """);

        await Assert.That(body.Descendants<Hyperlink>().Select(Text)).IsEquivalentTo(["Before"]);
    }

    const string headings =
        """


        # First

        # Second
        """;

    const string oneColumn =
        """
        | Contents |
        | --- |
        | [TOC] |
        """;

    // No column widths, so each column is taken to get an even share.
    const string twoColumns =
        """
        | Notes | Contents |
        | --- | --- |
        | Text | [TOC] |
        """;

    // Column widths in the ratio the grid is drawn in: a quarter and three quarters.
    const string weightedColumns =
        """
        +-----+---------------+
        | N   | Contents      |
        +=====+===============+
        | x   | [TOC]         |
        +-----+---------------+
        """;

    const string spanningCell =
        """
        +-----+-----+----------+
        | A   | B   | C        |
        +=====+=====+==========+
        | [TOC]     | x        |
        +-----+-----+----------+
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

    // The page number hangs off a right tab measured from the cell's own text edge, so the section's
    // width puts it past the cell's right edge by at least the cell margins.
    [Test]
    [Arguments(oneColumn, textWidth - cellMargins)]
    [Arguments(twoColumns, textWidth / 2 - cellMargins)]
    [Arguments(weightedColumns, textWidth * 3 / 4 - cellMargins)]
    [Arguments(spanningCell, textWidth / 2 - cellMargins)]
    [Arguments(cellInACell, textWidth - cellMargins - cellMargins)]
    public async Task InACellTheTabIsTheCellsInnerWidth(string table, int expected)
    {
        using var stream = await RenderStream(new StubResolver(), table + headings);
        using var doc = WordprocessingDocument.Open(stream, false);
        var entries = doc.MainDocumentPart!.Document!.Body!
            .Descendants<Paragraph>()
            .Where(_ => _.Elements<Hyperlink>().Any())
            .ToList();

        await Assert.That(entries.Count).IsEqualTo(2);
        await Assert.That(entries.All(_ => _.Ancestors<TableCell>().Any())).IsTrue();
        await Assert.That(string.Join(",", entries.Select(Tab))).IsEqualTo($"{expected},{expected}");

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    // A table in a list item is indented and sized to its content, so what it can grow to is the
    // text width less the indent.
    [Test]
    public async Task InAnIndentedTableTheTabAllowsForTheIndent()
    {
        var body = await Render(
            new StubResolver(),
            """
            - Item

              | Contents |
              | --- |
              | [TOC] |
            """ + headings);
        var indent = body.Descendants<TableIndentation>().Single().Width!.Value;
        var entries = body.Descendants<Paragraph>().Where(_ => _.Elements<Hyperlink>().Any());

        await Assert.That(indent).IsGreaterThan(0);
        await Assert.That(entries.Select(Tab)).IsEquivalentTo([textWidth - indent - cellMargins, textWidth - indent - cellMargins]);
    }

    // One in the body and one in a cell each list every heading, wherever it is, and each takes the
    // width of where it is itself.
    [Test]
    public async Task AFieldAndTheHeadingsItListsCanBeInDifferentContainers()
    {
        var body = await Render(
            new StubResolver(),
            $"[TOC]\n\n{oneColumn}\n\n# Before\n\n{headingInACell}\n");
        var entries = body.Descendants<Paragraph>().Where(_ => _.Elements<Hyperlink>().Any()).ToList();

        await Assert.That(entries.Select(_ => Text(_.Elements<Hyperlink>().Single())))
            .IsEquivalentTo(["Before", "Inside", "Before", "Inside"]);
        await Assert.That(entries.Select(Tab))
            .IsEquivalentTo([textWidth, textWidth, textWidth - cellMargins, textWidth - cellMargins]);

        // Both lists link to the one bookmark each heading has.
        await Assert.That(body.Descendants<Hyperlink>().Select(_ => _.Anchor?.Value).Distinct().Count()).IsEqualTo(2);
        await Assert.That(body.Descendants<BookmarkStart>().Count()).IsEqualTo(2);
    }

    // With nothing in range the field is never built, so there is nothing of it to resolve. Other
    // bookmarks resolving around it — the heading's own id here — must not be read as the field
    // being finished: cleared, it would show its placeholder for good.
    [Test]
    [Arguments("[TOC]")]
    [Arguments(oneColumn)]
    public async Task AFieldWithNoHeadingInRangeIsLeftForWord(string contents)
    {
        var body = await Render(
            new StubResolver(),
            contents +
            """


            #### Minor {#minor}

            See page [](#minor).
            """);
        var begins = body.Descendants<FieldChar>()
            .Where(_ => _.FieldCharType?.Value == FieldCharValues.Begin)
            .ToList();

        await Assert.That(body.Descendants<Hyperlink>().Any()).IsFalse();
        await Assert.That(begins.Count).IsEqualTo(2);

        // The table of contents is still asking; the page reference, which did resolve, is not.
        await Assert.That(begins[0].Dirty?.Value).IsTrue();
        await Assert.That(begins[1].Dirty?.Value).IsNull();
    }

    // Left for Word, nothing is built in either place: no entries, no bookmark on the heading in the
    // cell, and both fields still marked.
    [Test]
    public async Task WithoutAResolverBothAreLeftForWord()
    {
        var body = await Render(null, $"[TOC]\n\n{oneColumn}\n\n# Before\n\n{headingInACell}\n");
        var begins = body.Descendants<FieldChar>()
            .Where(_ => _.FieldCharType?.Value == FieldCharValues.Begin)
            .ToList();

        await Assert.That(body.Descendants<Hyperlink>().Any()).IsFalse();
        await Assert.That(body.Descendants<BookmarkStart>().Any()).IsFalse();
        await Assert.That(body.Descendants<TabStop>().Any()).IsFalse();
        await Assert.That(begins.Count).IsEqualTo(2);
        await Assert.That(begins.All(_ => _.Dirty?.Value == true)).IsTrue();
    }

    static async Task<MemoryStream> RenderStream(IPageNumberResolver? resolver, string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = resolver
        };
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        return stream;
    }

    static async Task<Body> Render(IPageNumberResolver? resolver, string markdown)
    {
        using var stream = await RenderStream(resolver, markdown);
        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    static int Tab(Paragraph entry) =>
        entry.ParagraphProperties!.GetFirstChild<Tabs>()!.Elements<TabStop>().Single().Position!.Value;

    static string Text(Hyperlink hyperlink) =>
        hyperlink.Elements<Run>().First().GetFirstChild<Text>()!.Text;

    static string PageNumber(Hyperlink hyperlink) =>
        hyperlink.Elements<Run>().Last().GetFirstChild<Text>()!.Text;
}
