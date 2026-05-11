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

    static Run RenderEmphasis(string markdown)
    {
        var paragraph = (ParagraphBlock)RendererHarness.Parse(markdown)[0];
        var emphasis = (EmphasisInline)paragraph.Inline!.FirstChild!;
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(emphasis);
        return (Run)renderer.Top.CurrentRuns[0];
    }
}
