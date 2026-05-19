public class CodeBlockRendererTests
{
    public class EmptyModel;


    [Test]
    public async Task EachLineBecomesCodeStyledParagraph()
    {
        var block = RendererHarness.FirstBlock<CodeBlock>("    line one\n    line two");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        var paragraphs = renderer.Drain().Cast<Paragraph>().ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(2);

        foreach (var p in paragraphs)
        {
            await Assert.That(p.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Code");
            var run = p.GetFirstChild<Run>()!;
            var fonts = run.RunProperties!.GetFirstChild<RunFonts>()!;
            await Assert.That(fonts.Ascii?.Value).IsEqualTo("Consolas");
        }

        var texts = paragraphs.Select(p => p.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text).ToList();
        await Assert.That(texts).IsEquivalentTo(["line one", "line two"]);
    }

    [Test]
    public async Task CodeBlockNestedInBlockQuoteIsIndented()
    {
        const string md =
            """
            > ```
            > line one
            > line two
            > ```
            """;

        var quote = RendererHarness.FirstBlock<QuoteBlock>(md);
        await Assert.That(quote.Descendants<CodeBlock>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(quote);

        var paragraphs = renderer.Drain()
            .OfType<Paragraph>()
            .Where(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Code")
            .ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        foreach (var p in paragraphs)
        {
            var indent = p.ParagraphProperties!.GetFirstChild<Indentation>()!;
            await Assert.That(indent.Left?.Value).IsEqualTo("720");
        }

        await VerifyDocument(md);
    }

    [Test]
    public async Task FencedCodeBlockWithLanguageProducesCodeStyledParagraphs()
    {
        const string md =
            """
            ```csharp
            var x = 1;
            var y = 2;
            ```
            """;

        var block = RendererHarness.FirstBlock<CodeBlock>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        var paragraphs = renderer.Drain().Cast<Paragraph>().ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        foreach (var p in paragraphs)
        {
            await Assert.That(p.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Code");
        }

        var texts = paragraphs.Select(p => p.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text).ToList();
        await Assert.That(texts).IsEquivalentTo(["var x = 1;", "var y = 2;"]);
    }

    [Test]
    public async Task EmptyFencedCodeBlockProducesNoParagraphs()
    {
        const string md =
            """
            ```
            ```
            """;

        var block = RendererHarness.FirstBlock<CodeBlock>(md);
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        await Assert.That(renderer.Drain().Count).IsEqualTo(0);
    }

    [Test]
    public async Task CodeBlockNestedInListItemIsIndented()
    {
        const string md =
            """
            - item

                  line one
                  line two
            """;

        var list = RendererHarness.FirstBlock<ListBlock>(md);
        await Assert.That(list.Descendants<CodeBlock>().Any()).IsTrue();

        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(list);

        var paragraphs = renderer.Drain()
            .OfType<Paragraph>()
            .Where(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Code")
            .ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        foreach (var p in paragraphs)
        {
            var indent = p.ParagraphProperties!.GetFirstChild<Indentation>()!;
            await Assert.That(indent.Left?.Value).IsEqualTo("480");
        }

        await VerifyDocument(md);
    }

    static async Task VerifyDocument(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("code", markdown, styleSource);
        using var stream = new MemoryStream();
        await store.Render("code", new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
