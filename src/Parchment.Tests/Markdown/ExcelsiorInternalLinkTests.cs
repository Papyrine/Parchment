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

    // The relationships are read once per table rather than searched per link, so what is left
    // in the part afterwards is checked link by link: every external link still resolves to its
    // own target, and no "#name" relationship survives.
    [Test]
    public async Task ManyRowsKeepEachExternalTargetAndDropEachInternalOne()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ContentsModel>("{{ Rows }}", styleSource);

        var rows = new List<ContentsRow>();
        for (var index = 0; index < 60; index++)
        {
            if (index % 2 == 0)
            {
                rows.Add(new() { Title = new($"https://example.com/{index}", $"External {index}"), Page = new($"#s{index}", "") });
                continue;
            }

            rows.Add(new() { Title = new($"#s{index}", $"Internal {index}"), Page = new($"https://example.com/page{index}", "Page") });
        }

        using var stream = new MemoryStream();
        await store.Render(new ContentsModel { Rows = rows }, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart!;
        var targets = mainPart.HyperlinkRelationships.ToDictionary(_ => _.Id, _ => _.Uri.OriginalString);
        var links = mainPart.Document!.Body!.Descendants<Hyperlink>().ToList();

        await Assert.That(targets.Count).IsEqualTo(60);
        await Assert.That(targets.Values.Any(_ => _.StartsWith('#'))).IsFalse();
        await Assert.That(links.Count).IsEqualTo(90);

        foreach (var link in links)
        {
            var text = link.InnerText;
            if (text.StartsWith("Internal ", StringComparison.Ordinal))
            {
                await Assert.That(link.Id).IsNull();
                await Assert.That(link.Anchor?.Value).IsEqualTo("s" + text["Internal ".Length..]);
                continue;
            }

            await Assert.That(link.Anchor).IsNull();
            var target = targets[link.Id!.Value!];
            if (text == "Page")
            {
                await Assert.That(target).StartsWith("https://example.com/page");
                continue;
            }

            await Assert.That(target).IsEqualTo("https://example.com/" + text["External ".Length..]);
        }

        var codes = mainPart.Document.Body.Descendants<FieldCode>().Select(_ => _.Text.Trim()).ToList();
        await Assert.That(codes.Count).IsEqualTo(30);
        await Assert.That(codes[0]).IsEqualTo("PAGEREF s0 \\h");
    }

    // Excelsior gives every link its own relationship, so this only arises for a table something
    // else has edited. The relationship is removed for the first hyperlink that uses it, and a
    // later one finds nothing to resolve and is left as it was, the same as a hyperlink whose id
    // never named a relationship.
    [Test]
    public async Task HyperlinksOutsideWhatExcelsiorProducesAreLeftAlone()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new(new Body());
        mainPart.AddHyperlinkRelationship(new("#elsewhere", UriKind.Relative), true, "rIdUnused");
        mainPart.AddHyperlinkRelationship(new("#alpha", UriKind.Relative), true, "rIdShared");
        mainPart.AddHyperlinkRelationship(new("https://example.com"), true, "rIdExternal");

        static Hyperlink Link(string text) =>
            new(new Run(new Text(text)));

        var first = Link("first");
        first.Id = "rIdShared";
        var second = Link("second");
        second.Id = "rIdShared";
        var external = Link("external");
        external.Id = "rIdExternal";
        var externalAgain = Link("external again");
        externalAgain.Id = "rIdExternal";
        var dangling = Link("dangling");
        dangling.Id = "rIdMissing";
        var anchored = Link("anchored");
        anchored.Anchor = "beta";

        var table = new Table(
            new TableRow(
                new TableCell(
                    new Paragraph(first, second, external, externalAgain, dangling, anchored))));

        ExcelsiorInternalLinks.Rewrite(table, mainPart, new());

        await Assert.That(first.Anchor?.Value).IsEqualTo("alpha");
        await Assert.That(first.Id).IsNull();
        await Assert.That(second.Anchor).IsNull();
        await Assert.That(second.Id?.Value).IsEqualTo("rIdShared");
        await Assert.That(external.Id?.Value).IsEqualTo("rIdExternal");
        await Assert.That(externalAgain.Id?.Value).IsEqualTo("rIdExternal");
        await Assert.That(dangling.Id?.Value).IsEqualTo("rIdMissing");
        await Assert.That(anchored.Anchor?.Value).IsEqualTo("beta");

        // A "#name" relationship no hyperlink in the table uses belongs to something else.
        await Assert.That(mainPart.HyperlinkRelationships.Select(_ => _.Id))
            .IsEquivalentTo(["rIdUnused", "rIdExternal"]);
    }

    class StubPageNumbers(Dictionary<string, int> pages) :
        IPageNumberResolver
    {
        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream document, Cancel cancel) =>
            Task.FromResult<IReadOnlyDictionary<string, int>>(pages);
    }
}
