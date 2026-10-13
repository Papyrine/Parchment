public class EmphasisInlineRendererTests
{
    [Test]
    public async Task SingleAsteriskEmitsItalic()
    {
        var run = RenderEmphasis("*text*");
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
        await Assert.That(run.RunProperties.GetFirstChild<Bold>()).IsNull();
    }

    [Test]
    public async Task DoubleAsteriskEmitsBold()
    {
        var run = RenderEmphasis("**text**");
        await Assert.That(run.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
    }

    [Test]
    public async Task SingleUnderscoreEmitsItalic()
    {
        var run = RenderEmphasis("_text_");
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
    }

    [Test]
    public async Task DoubleTildeEmitsStrike()
    {
        var run = RenderEmphasis("~~text~~");
        await Assert.That(run.RunProperties!.GetFirstChild<Strike>()).IsNotNull();
    }

    [Test]
    public async Task SingleTildeEmitsSubscript()
    {
        var run = RenderEmphasis("~text~");
        var alignment = run.RunProperties!.GetFirstChild<VerticalTextAlignment>()!;
        await Assert.That(alignment.Val?.Value).IsEqualTo(VerticalPositionValues.Subscript);
    }

    [Test]
    public async Task CaretEmitsSuperscript()
    {
        var run = RenderEmphasis("^text^");
        var alignment = run.RunProperties!.GetFirstChild<VerticalTextAlignment>()!;
        await Assert.That(alignment.Val?.Value).IsEqualTo(VerticalPositionValues.Superscript);
    }

    [Test]
    public async Task PlusEmitsUnderline()
    {
        var run = RenderEmphasis("++text++");
        await Assert.That(run.RunProperties!.GetFirstChild<Underline>()).IsNotNull();
    }

    [Test]
    public async Task EqualsEmitsHighlight()
    {
        var run = RenderEmphasis("==text==");
        var highlight = run.RunProperties!.GetFirstChild<Highlight>()!;
        await Assert.That(highlight.Val?.Value).IsEqualTo(HighlightColorValues.Yellow);
    }

    [Test]
    [Arguments("**[text](https://example.com)**")]
    [Arguments("**[text](#place)**")]
    [Arguments("**<https://example.com>**")]
    public async Task BoldAroundLinkReachesTheRunInsideTheHyperlink(string markdown)
    {
        var run = RenderLinkRuns(markdown).Single();
        await Assert.That(run.RunProperties!.Bold).IsNotNull();
        await Assert.That(run.RunProperties.RunStyle!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task ItalicAroundLinkReachesTheRunInsideTheHyperlink()
    {
        var run = RenderLinkRuns("*[text](https://example.com)*").Single();
        await Assert.That(run.RunProperties!.Italic).IsNotNull();
        await Assert.That(run.RunProperties.RunStyle!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task StrikeAroundLinkReachesTheRunInsideTheHyperlink()
    {
        var run = RenderLinkRuns("~~[text](https://example.com)~~").Single();
        await Assert.That(run.RunProperties!.Strike).IsNotNull();
        await Assert.That(run.RunProperties.RunStyle!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task NestedEmphasisAroundLinkAppliesBoth()
    {
        var run = RenderLinkRuns("***[text](https://example.com)***").Single();
        await Assert.That(run.RunProperties!.Bold).IsNotNull();
        await Assert.That(run.RunProperties.Italic).IsNotNull();
        await Assert.That(run.RunProperties.RunStyle!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task EmphasisAroundTextAndLinkReachesEveryRun()
    {
        var paragraph = RenderParagraph("**see [text](https://example.com) now**");
        var runs = paragraph.Descendants<Run>().ToList();
        await Assert.That(string.Join('|', runs.Select(_ => _.InnerText))).IsEqualTo("see |text| now");
        foreach (var run in runs)
        {
            await Assert.That(run.RunProperties!.Bold).IsNotNull();
        }

        await Assert.That(runs[0].RunProperties!.RunStyle).IsNull();
        await Assert.That(runs[1].RunProperties!.RunStyle!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task StyleOnEmphasisAroundLinkReplacesTheHyperlinkStyle()
    {
        var paragraph = RenderParagraph("**see [text](https://example.com)**{.Lead}");
        var runs = paragraph.Descendants<Run>().ToList();
        await Assert.That(runs[0].RunProperties!.RunStyle!.Val?.Value).IsEqualTo("Lead");
        await Assert.That(runs[1].RunProperties!.RunStyle!.Val?.Value).IsEqualTo("Lead");
        await Assert.That(runs[1].RunProperties!.Bold).IsNotNull();
    }

    static List<Run> RenderLinkRuns(string markdown) =>
        RenderParagraph(markdown).Descendants<Hyperlink>().Single().Elements<Run>().ToList();

    // Through the whole pipeline rather than a lone inline, so the paragraph can be checked
    // against the schema: rPr is a strict sequence, and a link run already carries an rStyle.
    static Paragraph RenderParagraph(string markdown)
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown(markdown).Single();
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(paragraph)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        if (errors.Count > 0)
        {
            throw new InvalidOperationException(string.Join(Environment.NewLine, errors));
        }

        return paragraph;
    }

    static Run RenderEmphasis(string markdown)
    {
        var paragraph = (ParagraphBlock)RendererHarness.Parse(markdown)[0];
        var emphasis = (EmphasisInline)paragraph.Inline!.FirstChild!;
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(emphasis);
        return (Run)renderer.Top.CurrentRuns[0];
    }
}
