public class HeadingBlockRendererTests
{
    [Test]
    public async Task Heading1ProducesHeading1Style()
    {
        var paragraph = RenderHeading("# Title");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading1");
    }

    [Test]
    public async Task Heading3ProducesHeading3Style()
    {
        var paragraph = RenderHeading("### Sub");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading3");
    }

    [Test]
    public async Task HeadingOffsetShiftsLevel()
    {
        var paragraph = RenderHeading("# Title", headingOffset: 2);
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading3");
    }

    [Test]
    public async Task HeadingOffsetClampsToNine()
    {
        var paragraph = RenderHeading("###### Six", headingOffset: 10);
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading9");
    }

    [Test]
    public async Task GenericAttributeOverridesStyle()
    {
        var paragraph = RenderHeading("# Title {.MyHeading}");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("MyHeading");
    }

    [Test]
    public async Task SetextLevel1ProducesHeading1Style()
    {
        var paragraph = RenderHeading("Title\n=====");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading1");
        await Assert.That(paragraph.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text).IsEqualTo("Title");
    }

    [Test]
    public async Task SetextLevel2ProducesHeading2Style()
    {
        var paragraph = RenderHeading("Title\n-----");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    [Test]
    public async Task HeadingWithItalicProducesItalicRun()
    {
        var paragraph = RenderHeading("# *italic*");
        var run = paragraph.GetFirstChild<Run>()!;
        await Assert.That(run.RunProperties!.GetFirstChild<Italic>()).IsNotNull();
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo("italic");
    }

    [Test]
    public async Task HeadingWithBoldProducesBoldRun()
    {
        var paragraph = RenderHeading("## **strong**");
        var run = paragraph.GetFirstChild<Run>()!;
        await Assert.That(run.RunProperties!.GetFirstChild<Bold>()).IsNotNull();
    }

    static Paragraph RenderHeading(string markdown, int headingOffset = 0)
    {
        var heading = RendererHarness.FirstBlock<HeadingBlock>(markdown);
        var renderer = RendererHarness.BuildRenderer(headingOffset);
        renderer.Render(heading);
        return (Paragraph)renderer.Drain().Single();
    }
}
