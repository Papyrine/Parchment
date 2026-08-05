public class HtmlBlockRendererTests
{
    public class EmptyModel;


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
