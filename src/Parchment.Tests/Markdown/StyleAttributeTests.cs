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

    // The renderer's own defaults stand down for a styled table, but an explicit column alignment
    // is authored intent rather than something the renderer invented, so it still applies. Only the
    // implicit header centre — which the author did not ask for — is suppressed.
    [Test]
    public async Task ExplicitColumnAlignmentSurvivesATableStyle()
    {
        var table = Render("{.BrandTable}\n| A | B |\n|:-:|---|\n| 1 | 2 |");
        var cells = table.Descendants<TableCell>().ToList();

        var aligned = cells[0].GetFirstChild<Paragraph>()!.ParagraphProperties!.GetFirstChild<Justification>();
        await Assert.That(aligned!.Val?.Value).IsEqualTo(JustificationValues.Center);

        // The unaligned header cell takes nothing, since the implicit centre stood down.
        await Assert.That(cells[1].GetFirstChild<Paragraph>()!.ParagraphProperties?.GetFirstChild<Justification>())
            .IsNull();
    }

    // A quote takes the attribute on the line above, which covers every line of it.
    [Test]
    public async Task QuoteStyleIsApplied()
    {
        // A blank quoted line makes two paragraphs; without it the lines lazily continue into one.
        var styles = RenderParagraphStyles<QuoteBlock>("{.PullQuote}\n> one\n>\n> two");
        await Assert.That(styles).IsEquivalentTo(["PullQuote", "PullQuote"]);
    }

    // Written at the end of a quoted line it binds to that line alone, and the narrower one wins.
    [Test]
    public async Task QuoteLineStyleBeatsTheBlockStyle()
    {
        var styles = RenderParagraphStyles<QuoteBlock>("{.PullQuote}\n> one {.Attribution}\n>\n> two");
        await Assert.That(styles).IsEquivalentTo(["Attribution", "PullQuote"]);
    }

    [Test]
    public async Task UnstyledQuoteKeepsTheDefault()
    {
        var styles = RenderParagraphStyles<QuoteBlock>("> one");
        await Assert.That(styles).IsEquivalentTo(["Quote"]);
    }

    [Test]
    public async Task CodeBlockStyleIsApplied()
    {
        var styles = RenderParagraphStyles<CodeBlock>("``` {.Snippet}\nvar x = 1;\n```");
        await Assert.That(styles).IsEquivalentTo(["Snippet"]);
    }

    // Markdig turns a fence's info string into a `language-xxx` class. It is synthesised, not
    // written, so treating it as a style would give every ```csharp block a style called
    // language-csharp that nobody defined.
    [Test]
    public async Task FenceLanguageIsNotMistakenForAStyle()
    {
        var styles = RenderParagraphStyles<CodeBlock>("```csharp\nvar x = 1;\n```");
        await Assert.That(styles).IsEquivalentTo(["Code"]);
    }

    [Test]
    public async Task CodeBlockStyleWinsOverTheFenceLanguage()
    {
        var styles = RenderParagraphStyles<CodeBlock>("```csharp {.Snippet}\nvar x = 1;\n```");
        await Assert.That(styles).IsEquivalentTo(["Snippet"]);
    }

    [Test]
    public async Task CodeInlineStyleBecomesRunStyle()
    {
        var run = FirstRun("Text `code`{.Mono} tail");
        await Assert.That(run.RunProperties!.RunStyle!.Val?.Value).IsEqualTo("Mono");
    }

    [Test]
    public async Task LinkStyleOverridesHyperlink()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ParagraphBlock>("[text](http://x){.Reference}"));
        var hyperlink = ((Paragraph) renderer.Drain()[0]).Descendants<Hyperlink>().Single();
        var run = hyperlink.Elements<Run>().First();

        await Assert.That(run.RunProperties!.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Reference");
    }

    [Test]
    public async Task UnstyledLinkKeepsHyperlink()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ParagraphBlock>("[text](http://x)"));
        var hyperlink = ((Paragraph) renderer.Drain()[0]).Descendants<Hyperlink>().Single();
        var run = hyperlink.Elements<Run>().First();

        await Assert.That(run.RunProperties!.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Hyperlink");
    }

    [Test]
    public async Task AutolinkStyleOverridesHyperlink()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ParagraphBlock>("<http://x>{.Reference}"));
        var hyperlink = ((Paragraph) renderer.Drain()[0]).Descendants<Hyperlink>().Single();
        var run = hyperlink.Elements<Run>().First();

        await Assert.That(run.RunProperties!.GetFirstChild<RunStyle>()!.Val?.Value).IsEqualTo("Reference");
    }

    static Run FirstRun(string markdown)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<ParagraphBlock>(markdown));
        return ((Paragraph) renderer.Drain()[0])
            .Elements<Run>()
            .First(_ => _.RunProperties?.RunStyle != null);
    }

    static List<string> RenderParagraphStyles<TBlock>(string markdown)
        where TBlock : Block
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstBlock<TBlock>(markdown));
        return renderer.Drain()
            .OfType<Paragraph>()
            .Select(_ => _.ParagraphProperties!.ParagraphStyleId!.Val!.Value!)
            .ToList();
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
