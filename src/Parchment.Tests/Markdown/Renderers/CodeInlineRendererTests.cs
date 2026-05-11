public class CodeInlineRendererTests
{
    [Test]
    public async Task EmitsRunWithConsolasFont()
    {
        var inline = RendererHarness.FirstInline<CodeInline>("`x = 1`");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(inline);

        var run = (Run)renderer.Top.CurrentRuns.Single();
        var fonts = run.RunProperties!.GetFirstChild<RunFonts>()!;
        await Assert.That(fonts.Ascii?.Value).IsEqualTo("Consolas");
        await Assert.That(fonts.HighAnsi?.Value).IsEqualTo("Consolas");

        var text = run.GetFirstChild<Text>()!;
        await Assert.That(text.Text).IsEqualTo("x = 1");
        await Assert.That(text.Space?.Value).IsEqualTo(SpaceProcessingModeValues.Preserve);
    }
}
