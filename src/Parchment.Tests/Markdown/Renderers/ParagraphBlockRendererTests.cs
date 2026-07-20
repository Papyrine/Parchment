public class ParagraphBlockRendererTests
{
    [Test]
    public async Task PlainParagraphHasNoStyle()
    {
        var paragraph = Render("just text");
        await Assert.That(paragraph.ParagraphProperties).IsNull();
        await Assert.That(paragraph.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text).IsEqualTo("just text");
    }

    [Test]
    public async Task GenericAttributeSetsStyle()
    {
        var paragraph = Render("{.Caption}\nCaption text");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Caption");
    }

    // The form the readme documents. Only the leading form had a test, so the documented one was
    // resting on nothing.
    [Test]
    public async Task TrailingGenericAttributeSetsStyle()
    {
        var paragraph = Render("Some intro paragraph. {.IntroBlock}");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("IntroBlock");
    }

    // A Word style is one name, so everything else the syntax can carry has nowhere to map. These
    // pin that the extras are ignored rather than breaking the style that is there.
    [Test]
    public async Task FirstClassWinsOverLaterOnes()
    {
        var paragraph = Render("Text {.Caption .Muted}");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Caption");
    }

    [Test]
    public async Task IdIsIgnored()
    {
        var paragraph = Render("Text {#intro .Caption}");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Caption");
    }

    [Test]
    public async Task KeyValuePropertyIsIgnored()
    {
        var paragraph = Render("Text {.Caption width=400}");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Caption");
    }

    // An id on its own names no style, so the paragraph stays unstyled rather than picking up the id.
    [Test]
    public async Task IdAloneLeavesTheParagraphUnstyled()
    {
        var paragraph = Render("Text {#intro}");
        await Assert.That(paragraph.ParagraphProperties).IsNull();
    }

    static Paragraph Render(string markdown)
    {
        var block = RendererHarness.FirstBlock<ParagraphBlock>(markdown);
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(block);
        return (Paragraph)renderer.Drain().Single();
    }
}
