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

    static Paragraph Render(string markdown)
    {
        var block = RendererHarness.FirstBlock<ParagraphBlock>(markdown);
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(block);
        return (Paragraph)renderer.Drain().Single();
    }
}
