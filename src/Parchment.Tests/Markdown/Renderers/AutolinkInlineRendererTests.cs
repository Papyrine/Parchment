public class AutolinkInlineRendererTests
{
    [Test]
    public async Task UrlAutolinkEmitsHyperlinkWithStyle()
    {
        var inline = new AutolinkInline("https://example.com") { IsEmail = false };
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(inline);

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        await Assert.That(string.IsNullOrEmpty(hyperlink.Id?.Value)).IsFalse();

        var run = hyperlink.GetFirstChild<Run>()!;
        await Assert.That(run.RunProperties!.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Hyperlink");
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("https://example.com");

        var rel = renderer.MainPart.HyperlinkRelationships.Single();
        await Assert.That(rel.Uri.ToString()).IsEqualTo("https://example.com/");
    }

    [Test]
    public async Task EmailAutolinkPrefixesMailto()
    {
        var inline = new AutolinkInline("user@example.com") { IsEmail = true };
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(inline);

        var rel = renderer.MainPart.HyperlinkRelationships.Single();
        await Assert.That(rel.Uri.ToString()).IsEqualTo("mailto:user@example.com");

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        var text = hyperlink.GetFirstChild<Run>()!.GetFirstChild<Text>()!;
        await Assert.That(text.Text).IsEqualTo("user@example.com");
    }
}
