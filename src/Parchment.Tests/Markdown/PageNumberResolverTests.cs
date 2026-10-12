using DocumentFormat.OpenXml.Validation;
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
    // entries, the styles, the links — is Parchment's to build.
    class StubResolver : IPageNumberResolver
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
                .Select((name, index) => (name: name!, page: index + 2))
                .ToDictionary(_ => _.name, _ => _.page, StringComparer.Ordinal);
            return Task.FromResult(pages);
        }
    }

    static async Task<MemoryStream> RenderStream(IPageNumberResolver? resolver)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = resolver
        };
        store.RegisterMarkdownTemplate<ReportModel>(markdown, styleSource);

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

    static async Task<Body> Render(IPageNumberResolver? resolver)
    {
        using var stream = await RenderStream(resolver);
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

    static string Text(Hyperlink hyperlink) =>
        hyperlink.Elements<Run>().First().GetFirstChild<Text>()!.Text;

    static string PageNumber(Hyperlink hyperlink) =>
        hyperlink.Elements<Run>().Last().GetFirstChild<Text>()!.Text;
}
