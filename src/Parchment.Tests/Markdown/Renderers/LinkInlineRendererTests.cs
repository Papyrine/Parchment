public class LinkInlineRendererTests
{
    [Test]
    public async Task EmitsHyperlinkWithRelIdAndStyle()
    {
        var link = RendererHarness.FirstInline<LinkInline>("[click](https://example.com)");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(link);

        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(1);
        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns[0];
        await Assert.That(string.IsNullOrEmpty(hyperlink.Id?.Value)).IsFalse();

        var run = hyperlink.GetFirstChild<Run>()!;
        var style = run.RunProperties!.GetFirstChild<RunStyle>()!;
        await Assert.That(style.Val?.Value).IsEqualTo("Hyperlink");

        var rel = renderer.MainPart.HyperlinkRelationships.Single();
        await Assert.That(rel.Uri.ToString()).IsEqualTo("https://example.com/");
    }

    [Test]
    public async Task EmptyUrlFallsThroughToChildren()
    {
        var link = new LinkInline { Url = string.Empty };
        link.AppendChild(new LiteralInline("plain"));
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(link);

        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(1);
        var run = (Run)renderer.Top.CurrentRuns[0];
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("plain");
    }
}
