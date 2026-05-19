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

    [Test]
    public async Task EmphasizedLinkTextCarriesBothItalicAndHyperlinkStyle()
    {
        var link = RendererHarness.FirstInline<LinkInline>("[*italic*](https://example.com)");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(link);

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        var run = hyperlink.GetFirstChild<Run>()!;
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
        await Assert.That(run.RunProperties.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Hyperlink");
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("italic");
    }

    [Test]
    public async Task BoldLinkTextCarriesBothBoldAndHyperlinkStyle()
    {
        var link = RendererHarness.FirstInline<LinkInline>("[**bold**](https://example.com)");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(link);

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        var run = hyperlink.GetFirstChild<Run>()!;
        await Assert.That(run.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
        await Assert.That(run.RunProperties.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task LinkWithNoChildrenProducesEmptyHyperlink()
    {
        var link = new LinkInline { Url = "https://example.com" };
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(link);

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        await Assert.That(hyperlink.ChildElements.Count).IsEqualTo(0);
        await Assert.That(string.IsNullOrEmpty(hyperlink.Id?.Value)).IsFalse();
    }

    [Test]
    public async Task MarkdownImageWithEmptyUrlFallsThroughToAltText()
    {
        var image = new LinkInline { Url = string.Empty, IsImage = true };
        image.AppendChild(new LiteralInline("alt"));
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(image);

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("alt");
    }
}
