public class ThematicBreakRendererTests
{
    public class EmptyModel;


    [Test]
    public async Task EmitsParagraphWithBottomBorder()
    {
        var block = RendererHarness.FirstBlock<ThematicBreakBlock>("---");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        var paragraph = (Paragraph)renderer.Drain().Single();
        var border = paragraph.ParagraphProperties!
            .GetFirstChild<ParagraphBorders>()!
            .GetFirstChild<BottomBorder>()!;
        await Assert.That(border.Val?.Value).IsEqualTo(BorderValues.Single);
        await Assert.That(border.Size?.Value).IsEqualTo((uint)6);
    }

    [Test]
    public async Task AsteriskVariantEmitsParagraphWithBottomBorder()
    {
        var block = RendererHarness.FirstBlock<ThematicBreakBlock>("***");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        var paragraph = (Paragraph)renderer.Drain().Single();
        var border = paragraph.ParagraphProperties!
            .GetFirstChild<ParagraphBorders>()!
            .GetFirstChild<BottomBorder>()!;
        await Assert.That(border.Val?.Value).IsEqualTo(BorderValues.Single);
    }

    [Test]
    public async Task UnderscoreVariantEmitsParagraphWithBottomBorder()
    {
        var block = RendererHarness.FirstBlock<ThematicBreakBlock>("___");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        var paragraph = (Paragraph)renderer.Drain().Single();
        var border = paragraph.ParagraphProperties!
            .GetFirstChild<ParagraphBorders>()!
            .GetFirstChild<BottomBorder>()!;
        await Assert.That(border.Val?.Value).IsEqualTo(BorderValues.Single);
    }

    [Test]
    public async Task ThematicBreakNestedInBlockQuoteIsIndented()
    {
        const string md = "> ---";

        var quote = RendererHarness.FirstBlock<QuoteBlock>(md);
        await Assert.That(quote.Descendants<ThematicBreakBlock>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(quote);

        var paragraph = renderer.Drain()
            .OfType<Paragraph>()
            .Single(_ => _.ParagraphProperties?.GetFirstChild<ParagraphBorders>() != null);
        var indent = paragraph.ParagraphProperties!.GetFirstChild<Indentation>()!;
        await Assert.That(indent.Left?.Value).IsEqualTo("720");

        await VerifyDocument(md);
    }

    [Test]
    public async Task ThematicBreakNestedInListItemIsIndented()
    {
        const string md =
            """
            - item

                ---
            """;

        var list = RendererHarness.FirstBlock<ListBlock>(md);
        await Assert.That(list.Descendants<ThematicBreakBlock>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(list);

        var paragraph = renderer.Drain()
            .OfType<Paragraph>()
            .Single(_ => _.ParagraphProperties?.GetFirstChild<ParagraphBorders>() != null);
        var indent = paragraph.ParagraphProperties!.GetFirstChild<Indentation>()!;
        await Assert.That(indent.Left?.Value).IsEqualTo("480");

        await VerifyDocument(md);
    }

    static async Task VerifyDocument(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("thematic", markdown, styleSource);
        using var stream = new MemoryStream();
        await store.Render("thematic", new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
