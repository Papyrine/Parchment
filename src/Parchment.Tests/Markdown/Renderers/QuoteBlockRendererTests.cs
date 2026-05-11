public class QuoteBlockRendererTests
{
    [Test]
    public async Task QuoteBlockProducesQuoteStyledParagraphs()
    {
        var quote = RendererHarness.FirstBlock<QuoteBlock>("> first line\n> second line");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(quote);

        var paragraphs = renderer.Drain().Cast<Paragraph>().ToList();
        await Assert.That(paragraphs.Count).IsGreaterThanOrEqualTo(1);
        foreach (var p in paragraphs)
        {
            await Assert.That(p.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Quote");
        }
    }
}
