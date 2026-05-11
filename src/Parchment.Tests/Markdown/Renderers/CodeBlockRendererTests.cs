public class CodeBlockRendererTests
{
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
}
