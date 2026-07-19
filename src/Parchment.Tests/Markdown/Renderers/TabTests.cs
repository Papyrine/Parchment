public class TabTests
{
    static Paragraph Render(string markdown)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ParagraphBlock>(markdown));
        return (Paragraph) renderer.Drain()[0];
    }

    // A tab left inside <w:t> is not a Word tab; Word renders it as ordinary whitespace. Nothing in
    // either flow reached <w:tab/> before, and html folds tabs to spaces, so a Word tab could not be
    // asked for at all.
    [Test]
    public async Task TabInTextBecomesWordTab()
    {
        var paragraph = Render("A\tSUBMISSIONS WITHIN AUTHORITY");
        var run = paragraph.GetFirstChild<Run>()!;

        var kinds = run.ChildElements.Select(_ => _.GetType().Name).ToList();
        await Assert.That(kinds).IsEquivalentTo(["Text", "TabChar", "Text"]);
        await Assert.That(run.Elements<Text>().First().Text).IsEqualTo("A");
        await Assert.That(run.Elements<Text>().Last().Text).IsEqualTo("SUBMISSIONS WITHIN AUTHORITY");
    }

    [Test]
    public async Task ConsecutiveTabsEachBecomeATab()
    {
        var run = Render("A\t\tB").GetFirstChild<Run>()!;

        var kinds = run.ChildElements.Select(_ => _.GetType().Name).ToList();
        await Assert.That(kinds).IsEquivalentTo(["Text", "TabChar", "TabChar", "Text"]);
    }

    // Markdown trims trailing whitespace from a line, so a tab at the end never reaches the
    // renderer. Put it between text to keep it.
    [Test]
    public async Task TrailingTabIsTrimmedByMarkdown()
    {
        var run = Render("A\t").GetFirstChild<Run>()!;

        var kinds = run.ChildElements.Select(_ => _.GetType().Name).ToList();
        await Assert.That(kinds).IsEquivalentTo(["Text"]);
    }

    [Test]
    public async Task TextWithoutTabsIsUnchanged()
    {
        var run = Render("plain text").GetFirstChild<Run>()!;

        var kinds = run.ChildElements.Select(_ => _.GetType().Name).ToList();
        await Assert.That(kinds).IsEquivalentTo(["Text"]);
    }
}
