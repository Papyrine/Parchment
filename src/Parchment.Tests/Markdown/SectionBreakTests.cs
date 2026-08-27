// ReSharper disable PartialTypeWithSinglePart
public partial class SectionBreakTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

    const string template =
        """
        Portrait to begin with.

        [SECTION]{orientation=landscape}

        A wide table belongs here.

        [SECTION]{orientation=portrait}

        And back to portrait.
        """;

    [Test]
    public async Task EachMarkerBecomesASection()
    {
        var sections = await Sections(template);

        // Two markers and the document-final section.
        await Assert.That(sections.Count).IsEqualTo(3);
        await Assert.That(Orientation(sections[0])).IsEqualTo("portrait");
        await Assert.That(Orientation(sections[1])).IsEqualTo("landscape");
        await Assert.That(Orientation(sections[2])).IsEqualTo("portrait");
    }

    // The marker reads "from here on", but a sectPr describes the section that ends at its
    // paragraph. So the first marker carries the settings it interrupted, not the ones it asked for.
    [Test]
    public async Task TheDeclaredSettingsLandOnTheSectionAfterTheMarker()
    {
        var sections = await Sections(template);
        var landscape = sections[1].GetFirstChild<PageSize>()!;

        await Assert.That(landscape.Width!.Value).IsEqualTo(8000u);
        await Assert.That(landscape.Height!.Value).IsEqualTo(6500u);
    }

    // Everything the style source set has to survive into every section, or a landscape page would
    // lose the headers, footers and page borders the rest of the document has.
    [Test]
    public async Task SectionsInheritTheStyleSource()
    {
        var sections = await Sections(template);

        foreach (var section in sections)
        {
            await Assert.That(section.GetFirstChild<PageBorders>()).IsNotNull();
            await Assert.That(section.GetFirstChild<PageMargin>()!.Header!.Value).IsEqualTo(720u);
        }
    }

    [Test]
    public async Task MarginsOverrideTheStyleSource()
    {
        var sections = await Sections(
            """
            Portrait.

            [SECTION]{orientation=landscape margins=100,200,300,400}

            Landscape.
            """);

        var margins = sections[1].GetFirstChild<PageMargin>()!;
        await Assert.That(margins.Top!.Value).IsEqualTo(100);
        await Assert.That(margins.Right!.Value).IsEqualTo(200u);
        await Assert.That(margins.Bottom!.Value).IsEqualTo(300);
        await Assert.That(margins.Left!.Value).IsEqualTo(400u);

        // margins= is the four page edges and nothing else. The header and footer distances have
        // attributes of their own, and a marker that does not name them leaves them alone.
        await Assert.That(margins.Header!.Value).IsEqualTo(720u);
        await Assert.That(margins.Footer!.Value).IsEqualTo(720u);
    }

    // The distance from the page edge to the header and footer, which is what decides how much room
    // is left for the text between them once the page edges have been pulled in.
    [Test]
    public async Task HeaderAndFooterOverrideTheStyleSource()
    {
        var sections = await Sections(
            """
            Portrait.

            [SECTION]{orientation=portrait margins=400,300,400,300 header=200 footer=150}

            Tighter.
            """);

        var margins = sections[1].GetFirstChild<PageMargin>()!;
        await Assert.That(margins.Header!.Value).IsEqualTo(200u);
        await Assert.That(margins.Footer!.Value).IsEqualTo(150u);
        await Assert.That(margins.Top!.Value).IsEqualTo(400);
    }

    // Either distance on its own, with no margins= beside it: a template moving the header up is not
    // obliged to restate the four page edges to do it.
    [Test]
    public async Task HeaderStandsAloneAndLeavesTheRestAsItWas()
    {
        var sections = await Sections(
            """
            Portrait.

            [SECTION]{orientation=portrait header=200}

            Header pulled up.
            """);

        var margins = sections[1].GetFirstChild<PageMargin>()!;
        await Assert.That(margins.Header!.Value).IsEqualTo(200u);
        await Assert.That(margins.Footer!.Value).IsEqualTo(720u);
        await Assert.That(margins.Top!.Value).IsEqualTo(500);
        await Assert.That(margins.Left!.Value).IsEqualTo(500u);
    }

    // Read one at a time, so a typo costs the setting it is in and nothing else. The alternative -
    // dropping the whole marker's page setup - would hide a working header behind a broken footer.
    [Test]
    public async Task AnUnparseableDistanceLeavesTheOthersStanding()
    {
        var sections = await Sections(
            """
            Portrait.

            [SECTION]{orientation=portrait header=200 footer=close}

            Header pulled up.
            """);

        var margins = sections[1].GetFirstChild<PageMargin>()!;
        await Assert.That(margins.Header!.Value).IsEqualTo(200u);
        await Assert.That(margins.Footer!.Value).IsEqualTo(720u);
    }

    // Word measures both distances from the page edge inwards, so there is nothing a negative one
    // could mean. Ignored like any other value that does not parse.
    [Test]
    public async Task ANegativeDistanceIsIgnored()
    {
        var sections = await Sections(
            """
            Portrait.

            [SECTION]{orientation=portrait header=-200}

            Unchanged.
            """);

        await Assert.That(sections[1].GetFirstChild<PageMargin>()!.Header!.Value).IsEqualTo(720u);
    }

    // Page numbering restarts wherever a section says it starts, so only the first section may keep
    // the style source's start. Otherwise every landscape page in a report would be page 1.
    [Test]
    public async Task OnlyTheFirstSectionRestartsPageNumbering()
    {
        using var styleSource = StyleSourceNumberedFromOne();
        var sections = await Sections(template, styleSource);

        await Assert.That(sections[0].GetFirstChild<PageNumberType>()!.Start!.Value).IsEqualTo(1);
        await Assert.That(sections[1].GetFirstChild<PageNumberType>()).IsNull();
        await Assert.That(sections[2].GetFirstChild<PageNumberType>()).IsNull();
    }

    // The page renders are the only assertion that shows the orientation actually changing: the
    // second page is wider than it is tall and the ones either side of it are not.
    [Test]
    public async Task SectionsRenderAtTheirOwnOrientation()
    {
        using var stream = await Render(template);
        await Verify(stream, "docx");
    }

    // An orientation the marker does not recognise is not a section break at all. Rendering it as
    // the literal text it already is makes a typo visible in the document, where a section break
    // that silently kept the current orientation would not be.
    [Test]
    public async Task UnknownOrientationStaysText()
    {
        var text = await Text("[SECTION]{orientation=sideways}");
        await Assert.That(text).Contains("[SECTION]");
        await Assert.That(await Sections("[SECTION]{orientation=sideways}")).Count().IsEqualTo(1);
    }

    [Test]
    public async Task MissingOrientationStaysText()
    {
        var text = await Text("[SECTION]");
        await Assert.That(text).Contains("[SECTION]");
        await Assert.That(await Sections("[SECTION]")).Count().IsEqualTo(1);
    }

    // Unparseable margins are ignored rather than rejected: the section break is still what the
    // author asked for, and the page keeps the margins it already had.
    [Test]
    public async Task UnparseableMarginsAreIgnored()
    {
        var sections = await Sections(
            """
            Portrait.

            [SECTION]{orientation=landscape margins=wide}

            Landscape.
            """);

        await Assert.That(sections).Count().IsEqualTo(2);
        await Assert.That(sections[1].GetFirstChild<PageMargin>()!.Top!.Value).IsEqualTo(500);
    }

    // Word has no way to express a section that starts inside a table cell, and markdown has no way
    // to ask for one: Markdig binds a generic attribute somewhere other than the cell, so the marker
    // never sees an orientation and stays the text it already was. Pinned because the alternative —
    // a sectPr inside a w:tc — is not a document Word can open, so it is worth knowing if a renderer
    // change ever makes the attribute reachable here. SectionBreaks.Apply rejects it either way.
    [Test]
    public async Task MarkerInsideATableStaysText()
    {
        const string markdown =
            """
            | One | Two |
            | --- | --- |
            | [SECTION]{orientation=landscape} | Cell |
            """;

        await Assert.That(await Text(markdown)).Contains("[SECTION]");
        await Assert.That(await Sections(markdown)).Count().IsEqualTo(1);
    }

    [Test]
    public async Task NoStyleSourceIsRejected()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("[SECTION]{orientation=landscape}");

        using var stream = new MemoryStream();
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(() => store.Render(new EmptyModel(), stream));

        await Assert.That(exception!.Message).Contains("style source");
    }

    static async Task<string> Text(string markdown)
    {
        using var stream = await Render(markdown);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }

    // Every section of the document, in order: the sectPr each section break carries, then the
    // document-final one.
    static async Task<IReadOnlyList<SectionProperties>> Sections(string markdown, Stream? styleSource = null)
    {
        using var stream = await Render(markdown, styleSource);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!
            .Descendants<SectionProperties>()
            .ToList();
    }

    static async Task<MemoryStream> Render(string markdown, Stream? styleSource = null)
    {
        using var built = styleSource == null ? DocxTemplateBuilder.Build() : null;
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource ?? built);

        var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        return stream;
    }

    static string Orientation(SectionProperties section)
    {
        var size = section.GetFirstChild<PageSize>()!;
        // Portrait is Word's assumption, so the attribute is only written for landscape.
        return size.Orient?.Value == PageOrientationValues.Landscape ? "landscape" : "portrait";
    }

    static MemoryStream StyleSourceNumberedFromOne()
    {
        var stream = DocxTemplateBuilder.Build();
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>()
                .Single()
                .AppendChild(
                    new PageNumberType
                    {
                        Start = 1
                    });
        }

        stream.Position = 0;
        return stream;
    }
}
