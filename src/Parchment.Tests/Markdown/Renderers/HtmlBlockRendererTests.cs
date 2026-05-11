public class HtmlBlockRendererTests
{
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
}
