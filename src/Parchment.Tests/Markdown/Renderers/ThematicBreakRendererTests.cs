public class ThematicBreakRendererTests
{
    [Test]
    public async Task EmitsParagraphWithBottomBorder()
    {
        var block = RendererHarness.FirstBlock<ThematicBreakBlock>("---");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(block);

        var paragraph = (Paragraph)renderer.Drain().Single();
        var border = paragraph.ParagraphProperties!
            .GetFirstChild<ParagraphBorders>()!
            .GetFirstChild<BottomBorder>()!;
        await Assert.That(border.Val?.Value).IsEqualTo(BorderValues.Single);
        await Assert.That(border.Size?.Value).IsEqualTo((uint)6);
    }
}
