public class AutolinkInlineRendererTests
{
    [Test]
    public async Task UrlAutolinkEmitsHyperlinkWithStyle()
    {
        var inline = new AutolinkInline("https://example.com") { IsEmail = false };
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(inline);

        var link = (Hyperlink)renderer.Top.CurrentRuns.Single();
        await Assert.That(string.IsNullOrEmpty(link.Id?.Value)).IsFalse();

        var run = link.GetFirstChild<Run>()!;
        await Assert.That(run.RunProperties!.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Hyperlink");
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("https://example.com");

        var rel = renderer.MainPart.HyperlinkRelationships.Single();
        await Assert.That(rel.Uri.ToString()).IsEqualTo("https://example.com/");
    }

    // What MarkdownEncoder wrote into a bound part of the url is undone, in the text and the target.
    [Test]
    public async Task EncoderEscapesAreUndone()
    {
        var inline = new AutolinkInline(@"https://example.com/A\-1\_x\.y") { IsEmail = false };
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(inline);

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        await Assert.That(hyperlink.InnerText).IsEqualTo("https://example.com/A-1_x.y");

        var rel = renderer.MainPart.HyperlinkRelationships.Single();
        await Assert.That(rel.Uri.OriginalString).IsEqualTo("https://example.com/A-1_x.y");
    }

    // A backslash is a path separator in a file: url, so \\ there is not an escaped backslash.
    [Test]
    public async Task FileUrlKeepsItsBackslashes()
    {
        var inline = new AutolinkInline(@"file:\\server\share\_drafts") { IsEmail = false };
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(inline);

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        await Assert.That(hyperlink.InnerText).IsEqualTo(@"file:\\server\share\_drafts");
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
