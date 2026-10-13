// ReSharper disable PartialTypeWithSinglePart

public partial class HtmlBlockRendererTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

    [Test]
    public async Task CommentBlockEmitsNothing()
    {
        var block = RendererHarness.FirstBlock<HtmlBlock>("<!-- a comment -->");
        await Assert.That(block.Type).IsEqualTo(HtmlBlockType.Comment);

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(block);

        await Assert.That(renderer.Drain().Count).IsEqualTo(0);
        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(0);
    }

    [Test]
    public async Task NonCommentBlockProducesElements()
    {
        var block = RendererHarness.FirstBlock<HtmlBlock>("<p>hello</p>");

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(block);

        await Assert.That(renderer.Drain().Count).IsGreaterThan(0);
    }

    [Test]
    public async Task HtmlBlockNestedInBlockQuoteIndentsParagraphs()
    {
        const string md = "> <p>hello</p>";

        var quote = RendererHarness.FirstBlock<QuoteBlock>(md);
        await Assert.That(quote.Descendants<HtmlBlock>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(quote);

        var paragraph = renderer.Drain().OfType<Paragraph>().First();
        var indent = paragraph.ParagraphProperties!.GetFirstChild<Indentation>()!;
        await Assert.That(indent.Left?.Value).IsEqualTo("720");

        await VerifyDocument(md);
    }

    [Test]
    public async Task HtmlBlockNestedInListItemIndentsParagraphs()
    {
        const string md =
            """
            - item

                <p>hello</p>
            """;

        var list = RendererHarness.FirstBlock<ListBlock>(md);
        await Assert.That(list.Descendants<HtmlBlock>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(list);

        var htmlParagraph = renderer.Drain()
            .OfType<Paragraph>()
            .Single(_ => _.ParagraphProperties?.NumberingProperties == null);
        var indent = htmlParagraph.ParagraphProperties!.GetFirstChild<Indentation>()!;
        await Assert.That(indent.Left?.Value).IsEqualTo("480");

        await VerifyDocument(md);
    }

    // The indent was appended to the converter's pPr or tblPr, landing after the w:jc, w:numPr
    // sibling order or table borders the html had already produced, which the schema rejects.
    [Test]
    [Arguments("> <p style=\"text-align:center\">centered</p>")]
    [Arguments("- a\n\n  <h2 style=\"text-align:right\">h</h2>")]
    [Arguments("> <ul><li>one</li></ul>")]
    [Arguments("- a\n\n  <ol><li>one</li></ol>")]
    [Arguments("- a\n\n  <table><tr><td>c</td></tr></table>")]
    [Arguments("> <table><tr><td>c</td></tr></table>")]
    [Arguments("> <table border=\"1\" style=\"width:50%\"><tr><td>c</td></tr></table>")]
    [Arguments("- a\n  - b\n\n    <table border=\"1\"><tr><td>c</td></tr></table>")]
    [Arguments("- <table><tr><td>c</td></tr></table>\n- next")]
    public async Task HtmlInsideAContainerIsSchemaValid(string md)
    {
        using var stream = await RenderDocument(md);
        using var doc = WordprocessingDocument.Open(stream, false);

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task HtmlTableInListItemIsIndentedUnderTheItem()
    {
        using var stream = await RenderDocument("- a\n\n  <table border=\"1\"><tr><td>c</td></tr></table>");
        using var doc = WordprocessingDocument.Open(stream, false);

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.ChildElements[0]).IsTypeOf<Paragraph>();
        var table = (Table) body.ChildElements[1];
        var properties = table.GetFirstChild<TableProperties>()!;
        await Assert.That(properties.TableIndentation!.Width?.Value).IsEqualTo(480);
        await Assert.That(properties.Elements<TableIndentation>().Count()).IsEqualTo(1);
    }

    // A direct w:ind replaces the left indent of the numbering level, so the container's indent is
    // added to the level's rather than standing in for it.
    [Test]
    public async Task HtmlListInBlockQuoteKeepsItsLevelIndent()
    {
        using var stream = await RenderDocument("> <ul><li>one</li></ul>");
        using var doc = WordprocessingDocument.Open(stream, false);

        var main = doc.MainDocumentPart!;
        var properties = main.Document!.Body!.Elements<Paragraph>().First().ParagraphProperties!;
        var numbering = properties.NumberingProperties!;
        var instance = main.NumberingDefinitionsPart!.Numbering!
            .Elements<NumberingInstance>()
            .Single(_ => _.NumberID?.Value == numbering.NumberingId!.Val!.Value);
        var level = main.NumberingDefinitionsPart.Numbering
            .Elements<AbstractNum>()
            .Single(_ => _.AbstractNumberId?.Value == instance.AbstractNumId!.Val!.Value)
            .Elements<Level>()
            .Single(_ => _.LevelIndex?.Value == (numbering.NumberingLevelReference?.Val?.Value ?? 0));
        var levelLeft = int.Parse(level.PreviousParagraphProperties!.Indentation!.Left!.Value!);

        await Assert.That(levelLeft).IsGreaterThan(0);
        await Assert.That(properties.Indentation!.Left?.Value).IsEqualTo((levelLeft + 720).ToString());
    }

    static async Task<MemoryStream> RenderDocument(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource);
        var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        return stream;
    }

    static async Task VerifyDocument(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource);
        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
