
// ReSharper disable PartialTypeWithSinglePart

// A Word paragraph has one style and markdown containers nest, so a block inside a container has
// to keep its own style and take the container as direct formatting.
public partial class NestedBlockStyleTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

    const string document =
        """
        [TOC]

        # Top

        - ## In a list item

          Item text.

        > ## In a quote
        >
        > Level one.
        >
        > > Level two.
        > >
        > > > Level three.
        """;

    [Test]
    public async Task HeadingInAListItemKeepsItsStyleAndTheMarker()
    {
        var paragraphs = Render("- ## Title\n- plain");

        var heading = paragraphs[0].ParagraphProperties!;
        await Assert.That(heading.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
        await Assert.That(heading.NumberingProperties!.NumberingLevelReference!.Val?.Value).IsEqualTo(0);
        await Assert.That(heading.NumberingProperties.NumberingId!.Val?.Value)
            .IsEqualTo(paragraphs[1].ParagraphProperties!.NumberingProperties!.NumberingId!.Val?.Value);
        await Assert.That(paragraphs[1].ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("ListParagraph");
    }

    [Test]
    public async Task HeadingInANestedListItemUsesTheNestingLevel()
    {
        var paragraphs = Render("- outer\n  - ### Inner");

        var heading = paragraphs[1].ParagraphProperties!;
        await Assert.That(heading.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading3");
        await Assert.That(heading.NumberingProperties!.NumberingLevelReference!.Val?.Value).IsEqualTo(1);
    }

    // Only an item's first block takes the marker, so a heading further down sits at the item's
    // text indent the way a continuation paragraph does.
    [Test]
    public async Task HeadingAfterAnItemsFirstParagraphIsIndentedNotNumbered()
    {
        var paragraphs = Render("- text\n\n  ## Title");

        var heading = paragraphs[1].ParagraphProperties!;
        await Assert.That(heading.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
        await Assert.That(heading.NumberingProperties).IsNull();
        await Assert.That(heading.Indentation!.Left?.Value).IsEqualTo("480");
    }

    [Test]
    public async Task HeadingInAListItemKeepsItsClassAndBookmark()
    {
        var paragraphs = Render("- ## Title {.MyHeading #anchor}");

        await Assert.That(paragraphs[0].ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("MyHeading");
        await Assert.That(paragraphs[0].Elements<BookmarkStart>().Single().Name?.Value).IsEqualTo("anchor");
    }

    [Test]
    public async Task HeadingInAQuoteKeepsItsStyleAtTheQuoteIndent()
    {
        var paragraphs = Render("> ## Title\n>\n> text");

        var heading = paragraphs[0].ParagraphProperties!;
        await Assert.That(heading.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
        await Assert.That(heading.Indentation!.Left?.Value).IsEqualTo("720");
        await Assert.That(paragraphs[1].ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Quote");
        await Assert.That(paragraphs[1].ParagraphProperties!.Indentation).IsNull();
    }

    [Test]
    public async Task NestedQuotesIndentByLevel()
    {
        var paragraphs = Render("> one\n>\n> > two\n> >\n> > > three\n> > >\n> > > ## Deep");

        var styles = paragraphs.Select(_ => _.ParagraphProperties!.ParagraphStyleId!.Val?.Value);
        await Assert.That(string.Join(' ', styles)).IsEqualTo("Quote Quote Quote Heading2");

        // The outermost level states nothing: the style's own indent is the first step.
        var indents = paragraphs.Select(_ => _.ParagraphProperties!.Indentation?.Left?.Value ?? "style");
        await Assert.That(string.Join(' ', indents)).IsEqualTo("style 1440 2160 2160");
    }

    [Test]
    public async Task QuoteStyleAttributeSurvivesNesting()
    {
        var paragraphs = Render("> one\n>\n> {.PullQuote}\n> > two");

        await Assert.That(paragraphs[1].ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("PullQuote");
        await Assert.That(paragraphs[1].ParagraphProperties!.Indentation!.Left?.Value).IsEqualTo("1440");
    }

    // The entries of a built table of contents come from the heading-styled paragraphs, so these
    // were missing from it as well as from Word's navigation pane.
    [Test]
    public async Task ContainedHeadingsAreListedAndTheDocumentIsValid()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = new StubResolver()
        };
        store.RegisterMarkdownTemplate<EmptyModel>(document, styleSource);
        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var entries = doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Where(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC", StringComparison.Ordinal) == true)
            .Select(_ => string.Concat(_.Descendants<Text>().Select(text => text.Text)))
            .ToList();
        await Assert.That(entries.Count).IsEqualTo(3);
        await Assert.That(entries[1]).StartsWith("In a list item");
        await Assert.That(entries[2]).StartsWith("In a quote");

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    class StubResolver : IPageNumberResolver
    {
        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel)
        {
            using var doc = WordprocessingDocument.Open(docx, false);
            IReadOnlyDictionary<string, int> pages = doc.MainDocumentPart!.Document!.Body!
                .Descendants<BookmarkStart>()
                .Select(_ => _.Name!.Value!)
                .Distinct()
                .ToDictionary(_ => _, _ => 1, StringComparer.Ordinal);
            return Task.FromResult(pages);
        }
    }

    static List<Paragraph> Render(string markdown) =>
        RendererHarness.RenderMarkdown(markdown).Cast<Paragraph>().ToList();
}
