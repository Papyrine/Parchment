using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Validation;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// Folding an id into a legal bookmark name maps different ids to the same name, and Word keeps one
/// bookmark per name. Every id therefore gets a name of its own, and a link is given the name of
/// the id it was written against.
/// </summary>
public partial class BookmarkNameTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required IReadOnlyList<string> Items { get; init; }

        [ExcelsiorTable]
        public required IReadOnlyList<Row> Rows { get; init; }
    }

    public class Row
    {
        [Column(Heading = "Section", Order = 1)]
        public required Link Title { get; init; }

        [Column(Order = 2)]
        public required Link Page { get; init; }
    }

    [Test]
    public async Task IdsThatFoldToOneNameGetDistinctNames()
    {
        var elements = RendererHarness.RenderMarkdown(
            """
            # One {#a-b}

            # Two {#a_b}

            # Three {#a.b}

            [one](#a-b) [two](#a_b) [three](#a.b) page [](#a.b)
            """);

        await Assert.That(Names(elements)).IsEquivalentTo(["a_b", "a_b_2", "a_b_3"]);
        await Assert.That(Anchors(elements)).IsEquivalentTo(["a_b", "a_b_2", "a_b_3"]);
        await Assert.That(elements[^1].InnerText).Contains("PAGEREF a_b_3 ");
    }

    // A link is routinely written above the heading it points at, so the name is settled at the
    // first mention of the id and the heading picks up what the link was given.
    [Test]
    public async Task LinkBeforeItsHeadingGetsTheHeadingsName()
    {
        var elements = RendererHarness.RenderMarkdown(
            """
            [two](#a_b) [one](#a-b)

            # One {#a-b}

            # Two {#a_b}
            """);

        await Assert.That(Anchors(elements)).IsEquivalentTo(["a_b", "a_b_2"]);
        await Assert.That(NameOf(elements[1])).IsEqualTo("a_b_2");
        await Assert.That(NameOf(elements[2])).IsEqualTo("a_b");
    }

    [Test]
    public async Task IdsThatDifferPastTheLimitGetDistinctNames()
    {
        var stem = new string('a', 50);
        var elements = RendererHarness.RenderMarkdown(
            $$"""
              # One {#{{stem}}-one}

              # Two {#{{stem}}-two}

              [two](#{{stem}}-two)
              """);

        var names = Names(elements);
        await Assert.That(names[0]).IsEqualTo(new('a', 40));
        await Assert.That(names[1]).IsEqualTo(new string('a', 38) + "_2");
        await Assert.That(Anchors(elements).Single()).IsEqualTo(names[1]);
    }

    // Word compares bookmark names without regard to case; html ids differ by it.
    [Test]
    public async Task IdsThatDifferOnlyByCaseGetDistinctNames()
    {
        var elements = RendererHarness.RenderMarkdown(
            """
            # One {#Intro}

            # Two {#intro}

            [two](#intro)
            """);

        await Assert.That(Names(elements)).IsEquivalentTo(["Intro", "intro_2"]);
        await Assert.That(Anchors(elements).Single()).IsEqualTo("intro_2");
    }

    [Test]
    [Arguments("überschrift-ß", "überschrift_ß", "überschrift_ß", "überschrift_ß_2")]
    [Arguments("日本-語", "日本.語", "日本_語", "日本_語_2")]
    [Arguments("1-a", "1_a", "_1_a", "_1_a_2")]
    public async Task NonAsciiAndLeadingDigitIdsAreKeptApart(string first, string second, string firstName, string secondName)
    {
        var elements = RendererHarness.RenderMarkdown(
            $$"""
              # One {#{{first}}}

              # Two {#{{second}}}

              [one](#{{first}}) [two](#{{second}})
              """);

        await Assert.That(Names(elements)).IsEquivalentTo([firstName, secondName]);
        await Assert.That(Anchors(elements)).IsEquivalentTo([firstName, secondName]);
    }

    // As in html, where the first element with an id is the one a link reaches. The repeat gets no
    // bookmark rather than a second one of the same name.
    [Test]
    public async Task RepeatedIdKeepsTheFirst()
    {
        var elements = RendererHarness.RenderMarkdown(
            """
            # One {#same}

            # Two {#same}

            Text {#same}

            [go](#same)
            """);

        await Assert.That(Names(elements)).IsEquivalentTo(["same"]);
        await Assert.That(NameOf(elements[0])).IsEqualTo("same");
        await Assert.That(elements.SelectMany(_ => _.Descendants<BookmarkEnd>()).Count()).IsEqualTo(1);
        await Assert.That(Anchors(elements).Single()).IsEqualTo("same");
    }

    // Nothing here can tell that no bookmark will carry the name: in the docx flow the target can be
    // a bookmark the template was authored with. The link is written as asked.
    [Test]
    public async Task LinkToAnIdNothingCarriesStillNamesIt()
    {
        var elements = RendererHarness.RenderMarkdown("[text](#missing-one) and [top](#)");

        await Assert.That(Anchors(elements)).IsEquivalentTo(["missing_one", "_"]);
    }

    [Test]
    public async Task EveryNameIsLegalAndUniqueAcrossTheDocument()
    {
        var resolver = new StubResolver();
        using var stream = await Render(resolver);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var starts = body.Descendants<BookmarkStart>().ToList();
        var names = starts.Select(_ => _.Name!.Value!).ToList();

        // Three explicit ids, the loop's heading once, and a generated name for each heading left
        // without one: the two repeats of the loop's heading and the two that never had an id.
        await Assert.That(names.Count).IsEqualTo(8);
        await Assert.That(names.Distinct(StringComparer.OrdinalIgnoreCase).Count()).IsEqualTo(8);
        foreach (var name in names)
        {
            await Assert.That(LegalName().IsMatch(name)).IsTrue();
        }

        // One table of contents entry per heading, each pointing at that heading's own bookmark.
        var headings = body.Elements<Paragraph>()
            .Where(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Heading1")
            .ToList();
        var entries = body.Elements<Paragraph>()
            .Where(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOC1")
            .ToList();
        await Assert.That(entries.Count).IsEqualTo(8);
        for (var index = 0; index < entries.Count; index++)
        {
            var anchor = entries[index].Descendants<Hyperlink>().Single().Anchor?.Value;
            await Assert.That(anchor).IsEqualTo(headings[index].Descendants<BookmarkStart>().Single().Name?.Value);
        }

        // The prose links, then the table's: each to the heading whose id it was written with.
        var bodyLinks = body.Elements<Paragraph>()
            .Where(_ => _.ParagraphProperties?.ParagraphStyleId == null)
            .SelectMany(_ => _.Elements<Hyperlink>())
            .ToList();
        await Assert.That(bodyLinks.Select(_ => Target(headings, _.Anchor!.Value!)))
            .IsEquivalentTo(["First", "Second", "Item"]);

        var table = body.Descendants<Table>().Single();
        await Assert.That(table.Descendants<Hyperlink>().Select(_ => Target(headings, _.Anchor!.Value!)))
            .IsEquivalentTo(["First", "Second"]);
        await Assert.That(table.InnerText).Contains($"PAGEREF {NameOf(headings[0])} ");
        await Assert.That(table.InnerText).Contains($"PAGEREF {NameOf(headings[1])} ");

        // The resolver is asked about the names the document carries.
        await Assert.That(resolver.Bookmarks.Distinct(StringComparer.OrdinalIgnoreCase).Count()).IsEqualTo(8);

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    [Test]
    public async Task NamesAreTheSameOnEveryRender()
    {
        using var first = await Render(new StubResolver());
        using var second = await Render(new StubResolver());

        await Assert.That(BodyXml(second)).IsEqualTo(BodyXml(first));
    }

    static string BodyXml(MemoryStream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.OuterXml;
    }

    // The headings' text is unique in the template below apart from the loop's, whose first is the
    // one that keeps the id.
    static string Target(List<Paragraph> headings, string anchor)
    {
        var heading = headings.First(_ => _.Descendants<BookmarkStart>().Any(start => start.Name?.Value == anchor));
        return string.Concat(heading.Descendants<Text>().Select(_ => _.Text));
    }

    // Three ids that fold to one name or collide with a generated one, a heading repeated by a
    // loop, and headings left for the table of contents to name.
    const string markdown =
        """
        [TOC]{levels=1}

        {{ Rows }}

        # First {#a-b}

        # Second {#a_b}

        # Reserved {#_Toc3}

        {% for item in Items %}
        # Item {#item}

        {{ item }}

        {% endfor %}

        # Same

        # Same

        [first](#a-b) [second](#a_b) [item](#item)
        """;

    static async Task<MemoryStream> Render(IPageNumberResolver resolver)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = resolver
        };
        store.RegisterMarkdownTemplate<Model>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(
            new Model
            {
                Items = ["one", "two", "three"],
                Rows =
                [
                    new() { Title = new("#a-b", "First"), Page = new("#a-b", "") },
                    new() { Title = new("#a_b", "Second"), Page = new("#a_b", "") }
                ]
            },
            stream);
        stream.Position = 0;
        return stream;
    }

    class StubResolver : IPageNumberResolver
    {
        public List<string> Bookmarks { get; } = [];

        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel)
        {
            using var doc = WordprocessingDocument.Open(docx, false);
            Bookmarks.AddRange(
                doc.MainDocumentPart!.Document!.Body!
                    .Descendants<BookmarkStart>()
                    .Select(_ => _.Name!.Value!));

            var pages = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var name in Bookmarks)
            {
                pages[name] = pages.Count + 2;
            }

            return Task.FromResult<IReadOnlyDictionary<string, int>>(pages);
        }
    }

    static List<string> Names(IReadOnlyList<OpenXmlElement> elements) =>
        elements
            .SelectMany(_ => _.Descendants<BookmarkStart>())
            .Select(_ => _.Name!.Value!)
            .ToList();

    static List<string> Anchors(IReadOnlyList<OpenXmlElement> elements) =>
        elements
            .SelectMany(_ => _.Descendants<Hyperlink>())
            .Select(_ => _.Anchor!.Value!)
            .ToList();

    static string? NameOf(OpenXmlElement element) =>
        element.Descendants<BookmarkStart>().SingleOrDefault()?.Name?.Value;

    [GeneratedRegex(@"^[\p{L}_][\p{L}\p{Nd}_]{0,39}$")]
    private static partial Regex LegalName();
}
