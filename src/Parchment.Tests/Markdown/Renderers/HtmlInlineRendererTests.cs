using Markdig.Syntax.Inlines;

public class HtmlInlineRendererTests
{
    [Test]
    public async Task EmptyTagEmitsNothing()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline(string.Empty));

        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(0);
    }

    [Test]
    public async Task BrTagEmitsBreakRun()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<br/>"));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.GetFirstChild<Break>()).IsNotNull();
    }

    [Test]
    public async Task EmOpenAndCloseAppliesItalicToInterveningRuns()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<em>"));
        renderer.AddRun(
            new Run(new Text("inside") { Space = SpaceProcessingModeValues.Preserve }));
        renderer.Render(new HtmlInline("</em>"));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.RunProperties).IsNotNull();
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
    }

    [Test]
    public async Task StrongOpenAndCloseAppliesBoldToInterveningRuns()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<strong>"));
        renderer.AddRun(
            new Run(new Text("inside") { Space = SpaceProcessingModeValues.Preserve }));
        renderer.Render(new HtmlInline("</strong>"));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
    }

    [Test]
    public async Task NestedTagsApplyBothFormats()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<em>"));
        renderer.Render(new HtmlInline("<strong>"));
        renderer.AddRun(
            new Run(new Text("inside") { Space = SpaceProcessingModeValues.Preserve }));
        renderer.Render(new HtmlInline("</strong>"));
        renderer.Render(new HtmlInline("</em>"));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
        await Assert.That(run.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
    }

    [Test]
    public async Task ClosingTagDoesNotAffectRunsEmittedAfter()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<em>"));
        renderer.AddRun(
            new Run(new Text("italic") { Space = SpaceProcessingModeValues.Preserve }));
        renderer.Render(new HtmlInline("</em>"));
        renderer.AddRun(
            new Run(new Text("plain") { Space = SpaceProcessingModeValues.Preserve }));

        var runs = renderer.Top.CurrentRuns.Cast<Run>().ToList();
        await Assert.That(runs[0].RunProperties!.GetFirstChild<Italic>()).IsNotNull();
        await Assert.That(runs[1].RunProperties).IsNull();
    }

    [Test]
    public async Task UnknownTagPassesThroughAsLiteralText()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<custom-thing>"));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("<custom-thing>");
    }
}
