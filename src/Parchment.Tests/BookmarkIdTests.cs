using DocumentFormat.OpenXml.Validation;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A bookmark's <c>w:id</c> has to be unique in the document. Everything that produces content
/// during a render — the html converter, the markdown renderer — numbers its own bookmarks from
/// 1, so each fragment is renumbered as it joins the document.
/// </summary>
public partial class BookmarkIdTests
{
    [AttributeUsage(AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        public required string Title;

        [Html]
        public required string Body;

        [Markdown]
        public required string Notes;
    }

    static Doc Model() =>
        new()
        {
            Title = "Report",
            Body = """<p id="x">hi</p><p id="y">yo</p>""",
            Notes = "## Notes {#notes}"
        };

    // The markdown is numbered when it is rendered, and the html is converted afterwards, when its
    // marker is swapped - so the converter's 1 and 2 used to land on the headings' 1 and 2.
    [Test]
    public async Task MarkdownFlowNumbersParkedHtmlAfterTheMarkdown()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("# Title {#top}\n\n{{ Body }}\n\n## Other {#other}\n");

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["top", "x", "y", "other"]);
        await AssertUniqueAndPaired(doc);
    }

    // Each conversion used to restart at 1 against the ids the template already had.
    [Test]
    public async Task DocxFlowNumbersProducedBookmarksAboveTheTemplates()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Body }}"),
                Paragraph("{{ Notes }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["existing", "x", "y", "notes"]);
        await AssertUniqueAndPaired(doc);

        // The template's own bookmark is left exactly as authored.
        var existing = Starts(doc).Single(_ => _.Name?.Value == "existing");
        await Assert.That(existing.Id?.Value).IsEqualTo("1");
    }

    // The anchors that mark token paragraphs are bookmarks too, numbered from the template's
    // highest id, and are stripped before the save by pairing an end to an anchor's start by id. A
    // produced bookmark that shared an anchor's id lost its end to that pass.
    [Test]
    public async Task DocxFlowProducedBookmarkKeepsItsEndBesideAnchoredParagraphs()
    {
        using var template = BuildDocx(
            [
                Paragraph("{{ Title }}"),
                Paragraph("{{ Body }}"),
                Paragraph("{{ Title }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["x", "y"]);
        await AssertUniqueAndPaired(doc);
    }

    // Ids are unique across the document rather than within a part, so what is produced into a
    // header is numbered from the same counter as the body.
    [Test]
    public async Task DocxFlowIdsAreUniqueAcrossParts()
    {
        using var template = BuildDocx(
            [Paragraph("{{ Body }}")],
            [
                Bookmarked("header bookmark", "inHeader", "7"),
                Paragraph("{{ Body }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["x", "y", "inHeader", "x", "y"]);
        await AssertUniqueAndPaired(doc);
    }

    // The markdown replaces the body, but a header arrives from the style source as authored.
    [Test]
    public async Task MarkdownFlowNumbersAboveTheStyleSourcesHeader()
    {
        using var styleSource = BuildDocx(
            [Bookmarked("replaced by the markdown", "gone", "9")],
            [Bookmarked("header bookmark", "inHeader", "1")]);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("# Title {#top}\n\n{{ Body }}\n", styleSource);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["top", "x", "y", "inHeader"]);
        await AssertUniqueAndPaired(doc);
    }

    [Test]
    public async Task IdsAreTheSameOnEveryRender()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Body }}"),
                Paragraph("{{ Notes }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var first = await Render(store, Model());
        using var second = await Render(store, Model());
        using var firstDoc = WordprocessingDocument.Open(first, false);
        using var secondDoc = WordprocessingDocument.Open(second, false);

        await Assert.That(Ids(secondDoc)).IsEquivalentTo(Ids(firstDoc));
    }

    static async Task AssertUniqueAndPaired(WordprocessingDocument doc)
    {
        var starts = Ids(doc);
        await Assert.That(starts.Distinct().Count()).IsEqualTo(starts.Count);

        // Every start is closed, once, in the part it opens in.
        foreach (var root in Roots(doc))
        {
            var opened = root.Descendants<BookmarkStart>().Select(_ => _.Id?.Value ?? "").Order().ToList();
            var closed = root.Descendants<BookmarkEnd>().Select(_ => _.Id?.Value ?? "").Order().ToList();
            await Assert.That(closed).IsEquivalentTo(opened);
        }

        var errors = new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static List<BookmarkStart> Starts(WordprocessingDocument doc) =>
        Roots(doc).SelectMany(_ => _.Descendants<BookmarkStart>()).ToList();

    static List<string> Names(WordprocessingDocument doc) =>
        Starts(doc).Select(_ => _.Name?.Value ?? "").ToList();

    static List<string> Ids(WordprocessingDocument doc) =>
        Starts(doc).Select(_ => _.Id?.Value ?? "").ToList();

    static IEnumerable<OpenXmlElement> Roots(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart!;
        yield return mainPart.Document!.Body!;
        foreach (var header in mainPart.HeaderParts)
        {
            yield return header.Header!;
        }
    }

    static async Task<MemoryStream> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildDocx(Paragraph[] body, Paragraph[]? header = null)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var documentBody = mainPart.Document.Body!;
            documentBody.Append(body);

            var section = new SectionProperties();
            if (header != null)
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
                headerPart.Header = new(header);
                section.Append(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    });
            }

            section.Append(
                new PageSize
                {
                    Width = 6500,
                    Height = 8000
                });
            documentBody.Append(section);
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph Paragraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    static Paragraph Bookmarked(string text, string name, string id) =>
        new(
            new BookmarkStart
            {
                Id = id,
                Name = name
            },
            new Run(new Text(text)),
            new BookmarkEnd
            {
                Id = id
            });
}
