public class PageBreakTests
{
    public class EmptyModel;

    static async Task<(string Text, int Breaks)> Render(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("t", markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render("t", new EmptyModel(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        return (body.InnerText, body.Descendants<PageBreakBefore>().Count());
    }

    // An empty page-break div needs no content to hang the css off — the paragraph carrying
    // PageBreakBefore is emitted from the element's own style.
    [Test]
    public async Task StandaloneEmptyDivEmitsPageBreak()
    {
        var result = await Render("""<div style="page-break-before: always"></div>""");
        await Assert.That(result.Breaks).IsEqualTo(1);
    }

    [Test]
    public async Task DivBetweenParagraphsEmitsPageBreak()
    {
        var result = await Render(
            """
            Some text

            <div style="page-break-before: always"></div>

            More text
            """);
        await Assert.That(result.Breaks).IsEqualTo(1);
        await Assert.That(result.Text).Contains("Some text");
        await Assert.That(result.Text).Contains("More text");
    }

    // Markdown only treats html as a block when it starts a line. Mid-line it is an inline, and an
    // inline div cannot carry a paragraph-level break — the tag has no paragraph of its own.
    [Test]
    public async Task DivMidLineIsInlineAndCannotBreak()
    {
        var result = await Render("""Text <div style="page-break-before: always"></div> more""");
        await Assert.That(result.Breaks).IsEqualTo(0);
    }
}
