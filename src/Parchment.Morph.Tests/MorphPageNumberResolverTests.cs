// The two halves meeting: Parchment writes the fields and the entries, Morph lays the document out
// and says which page each one landed on. Nothing here stubs the pagination.
public partial class MorphPageNumberResolverTests
{
    [ParchmentBindable]
    public partial class ReportModel
    {
        public required IReadOnlyList<string> Sections { get; init; }
    }

    const string Markdown =
        """
        Contents {.TOCHeading}

        [TOC]{levels=1}

        {% for section in Sections %}
        # {{ section }}

        Body text for this section.

        {% endfor %}
        """;

    static async Task<Body> Render(int sections)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = new MorphPageNumberResolver()
        };
        store.RegisterMarkdownTemplate<ReportModel>(Markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ReportModel
            {
                Sections = Enumerable.Range(1, sections).Select(_ => $"Section {_}").ToList()
            },
            stream);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    [Test]
    public async Task ShortDocumentPutsEveryEntryOnPageOne()
    {
        var body = await Render(3);

        await Assert.That(PageNumbers(body)).IsEquivalentTo(new List<string> { "1", "1", "1" });
    }

    // The real test of a resolver: entries that run past a page break report the page they are
    // actually on, which is the number Word would have computed.
    [Test]
    public async Task EntriesReportTheirOwnPageOnceTheDocumentSpansPages()
    {
        var body = await Render(60);
        var pages = PageNumbers(body).Select(int.Parse).ToList();

        await Assert.That(pages.Count).IsEqualTo(60);
        await Assert.That(pages[^1]).IsGreaterThan(pages[0]);
        // A contents list only makes sense if it climbs.
        await Assert.That(pages.SequenceEqual(pages.Order())).IsTrue();
    }

    // Sixty entries are pages of contents in their own right, and everything they point at sits
    // below them. Measuring before they existed would have reported page 1 for the first section
    // and been wrong by the length of the list it was about to insert.
    [Test]
    public async Task EntriesAccountForTheSpaceTheListItselfTakes()
    {
        var body = await Render(60);
        var pages = PageNumbers(body).Select(int.Parse).ToList();

        await Assert.That(pages[0]).IsGreaterThan(1);
    }

    [Test]
    public async Task NothingIsLeftForWordToRecalculate()
    {
        var body = await Render(60);

        await Assert.That(body.Descendants<FieldChar>().Any(_ => _.Dirty?.Value == true)).IsFalse();
    }

    static List<string> PageNumbers(Body body) =>
        body.Descendants<Hyperlink>()
            .Where(_ => _.Anchor?.Value?.StartsWith("_Toc") == true)
            .Select(_ => _.Elements<Run>().Last().GetFirstChild<Text>()!.Text)
            .ToList();
}
