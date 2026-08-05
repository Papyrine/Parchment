public class InternalLinkTests
{
    // A "#name" target is somewhere in this document, so it is an anchor rather than a relationship
    // to an external url — which is what it used to become, giving Word a link that went nowhere.
    [Test]
    public async Task AnchorUrlBecomesInternalHyperlink()
    {
        var elements = RendererHarness.RenderMarkdown("[Jump to the summary](#summary)");
        var hyperlink = ((Paragraph)elements.Single()).GetFirstChild<Hyperlink>()!;

        await Assert.That(hyperlink.Anchor?.Value).IsEqualTo("summary");
        await Assert.That(string.IsNullOrEmpty(hyperlink.Id?.Value)).IsTrue();
    }

    [Test]
    public async Task InternalLinkAddsNoExternalRelationship()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstInline<LinkInline>("[Jump](#summary)"));

        await Assert.That(renderer.MainPart.HyperlinkRelationships.Any()).IsFalse();
    }

    [Test]
    public async Task InternalLinkTextKeepsHyperlinkStyle()
    {
        var elements = RendererHarness.RenderMarkdown("[Jump](#summary)");
        var run = ((Paragraph)elements.Single()).GetFirstChild<Hyperlink>()!.GetFirstChild<Run>()!;

        await Assert.That(run.RunProperties!.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Hyperlink");
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("Jump");
    }

    // The link target is folded by the same rules as the bookmark it points at, so a name that
    // needed folding still resolves.
    [Test]
    public async Task AnchorIsFoldedLikeTheBookmarkName()
    {
        var elements = RendererHarness.RenderMarkdown("# Heading {#my-id}\n\n[Jump](#my-id)");
        var heading = (Paragraph)elements[0];
        var hyperlink = ((Paragraph)elements[1]).GetFirstChild<Hyperlink>()!;

        await Assert.That(hyperlink.Anchor?.Value).IsEqualTo(heading.GetFirstChild<BookmarkStart>()!.Name?.Value);
        await Assert.That(hyperlink.Anchor?.Value).IsEqualTo("my_id");
    }

    [Test]
    public async Task ExternalLinkStillUsesARelationship()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstInline<LinkInline>("[Site](https://example.com)"));

        var hyperlink = (Hyperlink)renderer.Top.CurrentRuns.Single();
        await Assert.That(string.IsNullOrEmpty(hyperlink.Id?.Value)).IsFalse();
        await Assert.That(hyperlink.Anchor?.Value).IsNull();
        await Assert.That(renderer.MainPart.HyperlinkRelationships.Count()).IsEqualTo(1);
    }

    // With no link text there is nothing to click, and the one useful thing left to say about a
    // place in the document is which page it is on.
    [Test]
    public async Task EmptyInternalLinkBecomesPageReference()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("Page [](#summary)").Single();

        await Assert.That(paragraph.Descendants<FieldCode>().Single().Text).IsEqualTo(" PAGEREF summary \\h ");
        await Assert.That(paragraph.Descendants<FieldChar>().First().Dirty?.Value).IsTrue();
    }

    [Test]
    public async Task PageReferenceCarriesAPlaceholderResult()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[](#summary)").Single();

        // Between the separate and end markers sits what a reader that never updates fields shows.
        var texts = paragraph.Descendants<Text>().Select(_ => _.Text).ToList();
        await Assert.That(texts).IsEquivalentTo(new List<string> { "1" });
    }

    [Test]
    public async Task EmptyExternalLinkIsStillAnEmptyHyperlink()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[](https://example.com)").Single();

        await Assert.That(paragraph.GetFirstChild<Hyperlink>()).IsNotNull();
        await Assert.That(paragraph.Descendants<FieldCode>().Any()).IsFalse();
    }
}
