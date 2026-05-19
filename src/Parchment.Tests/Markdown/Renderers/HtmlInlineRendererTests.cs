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
    public async Task ITagAppliesItalic()
    {
        var run = RenderWrapped("i");
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
    }

    [Test]
    public async Task BTagAppliesBold()
    {
        var run = RenderWrapped("b");
        await Assert.That(run.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
    }

    [Test]
    public async Task UTagAppliesSingleUnderline()
    {
        var run = RenderWrapped("u");
        var underline = run.RunProperties!.GetFirstChild<Underline>()!;
        await Assert.That(underline.Val?.Value).IsEqualTo(UnderlineValues.Single);
    }

    [Test]
    public async Task STagAppliesStrike()
    {
        var run = RenderWrapped("s");
        await Assert.That(run.RunProperties!.GetFirstChild<Strike>()).IsNotNull();
    }

    [Test]
    public async Task DelTagAppliesStrike()
    {
        var run = RenderWrapped("del");
        await Assert.That(run.RunProperties!.GetFirstChild<Strike>()).IsNotNull();
    }

    [Test]
    public async Task StrikeTagAppliesStrike()
    {
        var run = RenderWrapped("strike");
        await Assert.That(run.RunProperties!.GetFirstChild<Strike>()).IsNotNull();
    }

    [Test]
    public async Task SubTagAppliesSubscript()
    {
        var run = RenderWrapped("sub");
        var alignment = run.RunProperties!.GetFirstChild<VerticalTextAlignment>()!;
        await Assert.That(alignment.Val?.Value).IsEqualTo(VerticalPositionValues.Subscript);
    }

    [Test]
    public async Task SupTagAppliesSuperscript()
    {
        var run = RenderWrapped("sup");
        var alignment = run.RunProperties!.GetFirstChild<VerticalTextAlignment>()!;
        await Assert.That(alignment.Val?.Value).IsEqualTo(VerticalPositionValues.Superscript);
    }

    [Test]
    public async Task SelfClosingFormatterTagDoesNotApplyFormat()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<em/>"));
        renderer.AddRun(
            new Run(new Text("after") { Space = SpaceProcessingModeValues.Preserve }));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.RunProperties).IsNull();
    }

    [Test]
    public async Task TagWithAttributesStillAppliesFormat()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline("<em class=\"x\">"));
        renderer.AddRun(
            new Run(new Text("inside") { Space = SpaceProcessingModeValues.Preserve }));
        renderer.Render(new HtmlInline("</em>"));

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
    }

    static Run RenderWrapped(string tagName)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new HtmlInline($"<{tagName}>"));
        renderer.AddRun(
            new Run(new Text("inside") { Space = SpaceProcessingModeValues.Preserve }));
        renderer.Render(new HtmlInline($"</{tagName}>"));
        return (Run)renderer.Top.CurrentRuns.Single();
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
