
// ReSharper disable PartialTypeWithSinglePart

// The page number of a built entry hangs off a right tab at the text width, and the text width is a
// property of the section the table of contents is in — not of the document, once [SECTION] has
// given it more than one.
public partial class TableOfContentsSectionTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

    // DocxTemplateBuilder's page is 6500 by 8000 with 500 on every edge.
    const int portrait = 5500;
    const int landscape = 7000;

    // Only there so the entries are built at all. It places nothing, so they keep their placeholder.
    class StubResolver : IPageNumberResolver
    {
        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel) =>
            Task.FromResult<IReadOnlyDictionary<string, int>>(new Dictionary<string, int>());
    }

    [Test]
    public async Task WithoutASectionBreakTheStyleSourceSetsTheTab()
    {
        var tabs = await EntryTabs(
            """
            [TOC]

            # First

            # Second
            """);

        await Assert.That(tabs).IsEquivalentTo([portrait, portrait]);
    }

    // The reported case: the entries took the landscape width of the section the document ends in,
    // which on the portrait page they sit on is past the right margin.
    [Test]
    public async Task InTheFirstSectionTheTabIgnoresTheSectionsAfterIt()
    {
        var tabs = await EntryTabs(
            """
            [TOC]

            # First

            [SECTION]{orientation=landscape}

            # Second
            """);

        await Assert.That(tabs).IsEquivalentTo([portrait, portrait]);
    }

    // And its mirror: a portrait width on a landscape page stops the numbers short of the margin.
    [Test]
    public async Task InAMiddleSectionTheTabIsThatSections()
    {
        var tabs = await EntryTabs(
            """
            # Zero

            [SECTION]{orientation=landscape}

            [TOC]

            # First

            [SECTION]{orientation=portrait}

            # Second
            """);

        await Assert.That(tabs).IsEquivalentTo([landscape, landscape, landscape]);
    }

    [Test]
    public async Task InTheLastSectionTheTabIsTheDocumentFinalSections()
    {
        var tabs = await EntryTabs(
            """
            # First

            [SECTION]{orientation=landscape}

            [TOC]

            # Second
            """);

        await Assert.That(tabs).IsEquivalentTo([landscape, landscape]);
    }

    // margins= is top,right,bottom,left, so the landscape page's 8000 loses 200 and 400. The
    // section after it inherits those margins, which is why it cannot be told apart by them alone
    // and turns back to portrait.
    [Test]
    public async Task TheSectionsOwnMarginsAreTakenOff()
    {
        var tabs = await EntryTabs(
            """
            # First

            [SECTION]{orientation=landscape margins=100,200,300,400}

            [TOC]

            # Second

            [SECTION]{orientation=portrait}

            # Third
            """);

        await Assert.That(tabs).IsEquivalentTo([7400, 7400, 7400]);
    }

    [Test]
    public async Task EachTableOfContentsTakesItsOwnSection()
    {
        var tabs = await EntryTabs(
            """
            [TOC]{levels=1}

            # First

            [SECTION]{orientation=landscape}

            [TOC]{levels=1}

            # Second
            """);

        await Assert.That(tabs).IsEquivalentTo([portrait, portrait, landscape, landscape]);
    }

    // The entries are fitted to the cell, which takes its width from the section the table is in: a
    // one-column table is that section's text width less the 216 TableRenderer puts either side of a
    // cell's text. What this pins is which section is asked, found from a paragraph that is not itself
    // a child of the body.
    [Test]
    public async Task InsideATableTheTabIsTheTablesSections()
    {
        var tabs = await EntryTabs(
            """
            | Contents |
            | --- |
            | [TOC] |

            # First

            [SECTION]{orientation=landscape}

            # Second
            """);

        await Assert.That(tabs).IsEquivalentTo([portrait - 216, portrait - 216]);
    }

    [Test]
    public async Task OutputIsSchemaValid()
    {
        using var stream = await Render(
            """
            # Zero

            [SECTION]{orientation=landscape margins=100,200,300,400}

            [TOC]

            # First

            [SECTION]{orientation=portrait}

            # Second
            """);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    // The tab stop of every built entry, in document order.
    static async Task<List<int>> EntryTabs(string markdown)
    {
        using var stream = await Render(markdown);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!
            .Descendants<Paragraph>()
            .Where(_ => _.Elements<Hyperlink>().Any())
            .Select(_ => _.ParagraphProperties!.GetFirstChild<Tabs>()!.Elements<TabStop>().Single().Position!.Value)
            .ToList();
    }

    static async Task<MemoryStream> Render(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = new StubResolver()
        };
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        return stream;
    }
}
