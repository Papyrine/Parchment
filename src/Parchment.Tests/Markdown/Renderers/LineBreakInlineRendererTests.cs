public class LineBreakInlineRendererTests
{
    [Test]
    public async Task HardBreakEmitsBreakElement()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new LineBreakInline { IsHard = true });

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.GetFirstChild<Break>()).IsNotNull();
    }

    [Test]
    public async Task SoftBreakEmitsSpaceText()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new LineBreakInline { IsHard = false });

        var run = (Run)renderer.Top.CurrentRuns.Single();
        var text = run.GetFirstChild<Text>()!;
        await Assert.That(text.Text).IsEqualTo(" ");
        await Assert.That(text.Space?.Value).IsEqualTo(SpaceProcessingModeValues.Preserve);
    }
}
