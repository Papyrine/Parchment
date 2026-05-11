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

    static Paragraph RenderHeading(string markdown, int headingOffset = 0)
    {
        var heading = RendererHarness.FirstBlock<HeadingBlock>(markdown);
        var renderer = RendererHarness.BuildRenderer(headingOffset);
        renderer.Render(heading);
        return (Paragraph)renderer.Drain().Single();
    }
}
