public class LiteralInlineRendererTests
{
    [Test]
    public async Task EmitsTextRunWithPreserveSpace()
    {
        var literal = RendererHarness.FirstInline<LiteralInline>("hello world");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(literal);

        var runs = renderer.Top.CurrentRuns;
        await Assert.That(runs.Count).IsEqualTo(1);
        var run = (Run)runs[0];
        var text = run.GetFirstChild<Text>()!;
        await Assert.That(text.Text).IsEqualTo("hello world");
        await Assert.That(text.Space?.Value).IsEqualTo(SpaceProcessingModeValues.Preserve);
    }

    [Test]
    public async Task EmptyLiteralEmitsNothing()
    {
        var renderer = RendererHarness.BuildRenderer();
        var literal = new LiteralInline(string.Empty);

        renderer.Render(literal);

        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(0);
    }
}
