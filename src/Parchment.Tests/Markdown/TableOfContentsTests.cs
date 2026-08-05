public class TableOfContentsTests
{
    [Test]
    public async Task MarkerBecomesTocField()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[TOC]").Single();

        await Assert.That(Instruction(paragraph)).IsEqualTo(" TOC \\o \"1-3\" \\h \\z \\u ");
    }

    // Word offers to update a dirty field on open, which is the only way the entries and their page
    // numbers ever appear — layout is the renderer's to compute, not ours.
    [Test]
    public async Task FieldIsMarkedDirtySoWordBuildsIt()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[TOC]").Single();

        var begin = paragraph.Descendants<FieldChar>().First();
        await Assert.That(begin.FieldCharType?.Value).IsEqualTo(FieldCharValues.Begin);
        await Assert.That(begin.Dirty?.Value).IsTrue();
    }

    [Test]
    public async Task LevelsPropertySetsDepth()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[TOC]{levels=1}").Single();

        await Assert.That(Instruction(paragraph)).IsEqualTo(" TOC \\o \"1-1\" \\h \\z \\u ");
    }

    [Test]
    public async Task LevelsOutsideWordsRangeAreClamped()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[TOC]{levels=99}").Single();

        await Assert.That(Instruction(paragraph)).IsEqualTo(" TOC \\o \"1-9\" \\h \\z \\u ");
    }

    [Test]
    public async Task StyleAttributeAppliesToTheFieldParagraph()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[TOC]{.TOC1}").Single();

        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("TOC1");
    }

    // Only a paragraph that is nothing but the marker is a TOC request; the same text alongside
    // other content is what the author wrote and stays text.
    [Test]
    public async Task MarkerWithSurroundingTextStaysText()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("See [TOC] below").Single();

        await Assert.That(paragraph.Descendants<FieldChar>().Any()).IsFalse();
        await Assert.That(ParagraphText(paragraph)).IsEqualTo("See [TOC] below");
    }

    [Test]
    public async Task RealLinkNamedTocIsStillALink()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[TOC](https://example.com)").Single();

        await Assert.That(paragraph.GetFirstChild<Hyperlink>()).IsNotNull();
        await Assert.That(paragraph.Descendants<FieldChar>().Any()).IsFalse();
    }

    static string Instruction(Paragraph paragraph) =>
        paragraph.Descendants<FieldCode>().Single().Text;

    static string ParagraphText(Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<Text>().Select(_ => _.Text));
}
