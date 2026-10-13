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
        var link = new LinkInline
        {
            Url = "https://example.com"
        };
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

    // The link's character style goes onto runs its children already gave properties, and rStyle
    // has to lead the rPr sequence: Word throws away an element it finds out of place, which cost
    // the link its colour and underline exactly when its text was emphasised.
    [Test]
    [Arguments("[**bold** text](https://example.com)")]
    [Arguments("[*italic*](https://example.com)")]
    [Arguments("[`code`](https://example.com)")]
    [Arguments("[~~struck~~ and ***both***](https://example.com)")]
    [Arguments("[<u>underlined</u>](https://example.com)")]
    [Arguments("[**bold** text](#summary)")]
    [Arguments("[`code`](#summary)")]
    [Arguments("<https://example.com>")]
    public async Task RunStyleLeadsTheRunPropertiesOfEveryLinkRun(string markdown)
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown(markdown).Single();
        var runs = paragraph.Descendants<Hyperlink>().Single().Elements<Run>().ToList();

        await Assert.That(runs.Count).IsGreaterThan(0);
        foreach (var run in runs)
        {
            await Assert.That(run.RunProperties!.FirstChild).IsTypeOf<RunStyle>();
            await Assert.That(run.RunProperties.Elements<RunStyle>().Single().Val?.Value).IsEqualTo("Hyperlink");
        }

        await Assert.That(Validate(paragraph)).IsEmpty();
    }

    [Test]
    public async Task FormattedLinkTextKeepsItsFormattingBesideTheStyle()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[**bold**](https://example.com) and [`code`](https://example.com)").Single();
        var links = paragraph.Elements<Hyperlink>().ToList();

        await Assert.That(links[0].GetFirstChild<Run>()!.RunProperties!.Bold).IsNotNull();
        await Assert.That(links[1].GetFirstChild<Run>()!.RunProperties!.RunFonts!.Ascii?.Value).IsEqualTo("Consolas");
    }

    // A run can carry one character style. One a child asked for by name is more specific than the
    // link's, so it is kept rather than replaced or given a second rStyle beside it.
    [Test]
    [Arguments("[`code`{.Mono}](https://example.com)", "Mono")]
    [Arguments("[*lead*{.Lead}](https://example.com)", "Lead")]
    [Arguments("[<span class=\"Note\">text</span>](https://example.com)", "Note")]
    [Arguments("[`code`{.Mono}](https://example.com){.Reference}", "Mono")]
    public async Task ExplicitChildStyleWinsOverTheLinkStyle(string markdown, string expected)
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown(markdown).Single();
        var run = paragraph.Descendants<Hyperlink>().Single().Elements<Run>().Single();

        await Assert.That(run.RunProperties!.FirstChild).IsTypeOf<RunStyle>();
        await Assert.That(run.RunProperties.Elements<RunStyle>().Single().Val?.Value).IsEqualTo(expected);
        await Assert.That(Validate(paragraph)).IsEmpty();
    }

    [Test]
    public async Task UnstyledSiblingOfAStyledChildStillGetsTheLinkStyle()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[see `code`{.Mono}](https://example.com)").Single();
        var styles = paragraph.Descendants<Hyperlink>().Single()
            .Elements<Run>()
            .Select(_ => _.RunProperties!.RunStyle!.Val!.Value!)
            .ToList();

        await Assert.That(styles).IsEquivalentTo(["Hyperlink", "Mono"]);
    }

    static IEnumerable<string> Validate(OpenXmlElement element) =>
        new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(element)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}");
}
