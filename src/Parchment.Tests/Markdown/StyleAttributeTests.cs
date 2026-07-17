using MarkdigTable = Markdig.Extensions.Tables.Table;

public class StyleAttributeTests
{
    // {.StyleName} leads the table, on its own line, matching the heading and paragraph form. There
    // was previously no way to put a table style on a markdown table at all, and the renderer's own
    // direct formatting would have beaten one applied afterwards anyway.
    [Test]
    public async Task TableStyleIsApplied()
    {
        var table = Render("{.BrandTable}\n| A | B |\n|---|---|\n| 1 | 2 |");
        var properties = table.GetFirstChild<TableProperties>()!;

        await Assert.That(properties.GetFirstChild<TableStyle>()!.Val?.Value).IsEqualTo("BrandTable");
    }

    // Direct formatting beats a table style in Word, so the defaults have to stand down or the
    // style is silently overridden — the reason every branded table needed a post-render pass.
    [Test]
    public async Task TableStyleSuppressesHardcodedDirectFormatting()
    {
        var styled = Render("{.BrandTable}\n| A | B |\n|---|---|\n| 1 | 2 |");
        var styledProperties = styled.GetFirstChild<TableProperties>()!;

        await Assert.That(styledProperties.GetFirstChild<TableBorders>()).IsNull();
        await Assert.That(styledProperties.GetFirstChild<TableCellMarginDefault>()).IsNull();

        var headerRun = styled.Elements<TableRow>().First()
            .Elements<TableCell>().First()
            .GetFirstChild<Paragraph>()!;
        await Assert.That(headerRun.GetFirstChild<Run>()?.RunProperties?.GetFirstChild<Bold>()).IsNull();
        await Assert.That(headerRun.ParagraphProperties?.GetFirstChild<Justification>()).IsNull();
    }

    // Without a style the defaults are the only thing making the table legible, so they stay.
    [Test]
    public async Task TableWithoutStyleKeepsDefaultFormatting()
    {
        var plain = Render("| A | B |\n|---|---|\n| 1 | 2 |");
        var properties = plain.GetFirstChild<TableProperties>()!;

        await Assert.That(properties.GetFirstChild<TableStyle>()).IsNull();
        await Assert.That(properties.GetFirstChild<TableBorders>()).IsNotNull();
        await Assert.That(properties.GetFirstChild<TableCellMarginDefault>()).IsNotNull();
    }

    // The header still repeats across pages; only its formatting defers to the style.
    [Test]
    public async Task TableStyleStillRepeatsHeaderRow()
    {
        var styled = Render("{.BrandTable}\n| A | B |\n|---|---|\n| 1 | 2 |");
        var header = styled.Elements<TableRow>().First();

        await Assert.That(header.GetFirstChild<TableRowProperties>()!.GetFirstChild<TableHeader>()).IsNotNull();
    }

    // A list item's paragraph used to be hardcoded to ListParagraph, and it never delegated to the
    // renderer that reads attributes, so the item's own style could not reach it.
    [Test]
    public async Task ListItemStyleIsApplied()
    {
        var styles = RenderListStyles("- one{.ItemStyle}\n- two");
        await Assert.That(styles).IsEquivalentTo(["ItemStyle", "ListParagraph"]);
    }

    [Test]
    public async Task ListStyleAppliesToEveryItem()
    {
        var styles = RenderListStyles("{.WholeList}\n- one\n- two");
        await Assert.That(styles).IsEquivalentTo(["WholeList", "WholeList"]);
    }

    // A style on the item is more specific than one on the list.
    [Test]
    public async Task ItemStyleBeatsListStyle()
    {
        var styles = RenderListStyles("{.WholeList}\n- one{.ItemStyle}\n- two");
        await Assert.That(styles).IsEquivalentTo(["ItemStyle", "WholeList"]);
    }

    [Test]
    public async Task ListWithoutStyleKeepsListParagraph()
    {
        var styles = RenderListStyles("- one\n- two");
        await Assert.That(styles).IsEquivalentTo(["ListParagraph", "ListParagraph"]);
    }

    // The attribute binds to the emphasis inline, which nothing read — so the style was dropped
    // without a word while the emphasis itself still applied.
    [Test]
    public async Task EmphasisStyleBecomesRunStyle()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ParagraphBlock>("**text**{.Lead} tail"));
        var paragraph = (Paragraph) renderer.Drain()[0];
        var first = paragraph.Elements<Run>().First();

        await Assert.That(first.RunProperties!.RunStyle!.Val?.Value).IsEqualTo("Lead");

        // The emphasis still applies alongside the style.
        await Assert.That(first.RunProperties.GetFirstChild<Bold>()).IsNotNull();

        // And it does not leak onto the text after it.
        var last = paragraph.Elements<Run>().Last();
        await Assert.That(last.RunProperties?.RunStyle).IsNull();
    }

    static List<string> RenderListStyles(string markdown)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ListBlock>(markdown));
        return renderer.Drain()
            .OfType<Paragraph>()
            .Select(_ => _.ParagraphProperties!.ParagraphStyleId!.Val!.Value!)
            .ToList();
    }

    static Table Render(string markdown)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<MarkdigTable>(markdown));
        return (Table) renderer.Drain().Single();
    }
}
