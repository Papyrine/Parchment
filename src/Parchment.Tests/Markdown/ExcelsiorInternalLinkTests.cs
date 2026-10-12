// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A <c>#name</c> link in an Excelsior table means what it means everywhere else in the document:
/// a jump to a bookmark, or — with no text to click — the page that bookmark is on.
/// </summary>
public partial class ExcelsiorInternalLinkTests
{
    [ParchmentBindable]
    public partial class ContentsModel
    {
        [ExcelsiorTable]
        public required IReadOnlyList<ContentsRow> Rows { get; init; }
    }

    public class ContentsRow
    {
        [Column(Heading = "Section", Order = 1)]
        public required Link Title { get; init; }

        // No display text, so there is nothing to click and the only thing left to say about the
        // section is which page it starts on.
        [Column(Order = 2)]
        public required Link Page { get; init; }
    }

    const string markdown =
        """
        # Contents

        {{ Rows }}

        # Alpha {#alpha}

        Alpha body.

        # Beta {#beta}

        Beta body.
        """;

    static ContentsModel Contents() =>
        new()
        {
            Rows =
            [
                new() { Title = new("#alpha", "Alpha"), Page = new("#alpha", "") },
                new() { Title = new("#beta", "Beta"), Page = new("#beta", "") }
            ]
        };

    static async Task<MemoryStream> Render()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ContentsModel>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(Contents(), stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task TitleCellLinksToTheBookmark()
    {
        using var stream = await Render();
        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        var links = table.Descendants<Hyperlink>().ToList();

        await Assert.That(links.Count).IsEqualTo(2);
        await Assert.That(links.Select(_ => _.Anchor?.Value ?? "")).IsEquivalentTo(["alpha", "beta"]);

        // An anchored hyperlink carries no relationship, and leaving one behind is what makes Word
        // warn about the file on open.
        await Assert.That(links.All(_ => _.Id == null)).IsTrue();
    }

    [Test]
    public async Task EmptyLinkCellBecomesAPageReference()
    {
        using var stream = await Render();
        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        var codes = table.Descendants<FieldCode>().Select(_ => _.Text.Trim()).ToList();

        await Assert.That(codes).IsEquivalentTo(["PAGEREF alpha \\h", "PAGEREF beta \\h"]);
    }

    // The whole point of the rewrite: the fields the resolver fills have to be the ones the table
    // produced, or the page column stays on its placeholder.
    [Test]
    public async Task PageReferencesResolveToRealPageNumbers()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = new StubPageNumbers(new()
            {
                ["alpha"] = 4,
                ["beta"] = 9
            })
        };
        store.RegisterMarkdownTemplate<ContentsModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(Contents(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Descendants<Table>().Single();
        var results = table.Descendants<FieldChar>()
            .Where(_ => _.FieldCharType?.Value == FieldCharValues.Separate)
            .Select(_ => _.Ancestors<Paragraph>().First())
            .SelectMany(_ => _.Descendants<Text>().Select(text => text.Text))
            .ToList();

        await Assert.That(results).Contains("4");
        await Assert.That(results).Contains("9");
    }

    // An external link is left alone: only "#name" addresses this document.
    [Test]
    public async Task ExternalLinksKeepTheirRelationship()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ContentsModel>("{{ Rows }}", styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ContentsModel
            {
                Rows = [new() { Title = new("https://example.com", "Example"), Page = new("#alpha", "") }]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart!;
        var link = mainPart.Document!.Body!.Descendants<Hyperlink>().Single();

        await Assert.That(link.Anchor?.Value).IsNull();
        await Assert.That(link.Id?.Value).IsNotNull();
        await Assert.That(mainPart.HyperlinkRelationships.Single().Uri.OriginalString)
            .IsEqualTo("https://example.com");
    }

    class StubPageNumbers(Dictionary<string, int> pages) :
        IPageNumberResolver
    {
        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream document, Cancel cancel) =>
            Task.FromResult<IReadOnlyDictionary<string, int>>(pages);
    }
}
