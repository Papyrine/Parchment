public class ListBlockRendererTests
{
    public class EmptyModel;

    [Test]
    public async Task BulletListProducesListParagraphsWithNumbering()
    {
        const string md = "- one\n- two\n- three";
        var paragraphs = RenderList(md);

        await Assert.That(paragraphs.Count).IsEqualTo(3);
        foreach (var p in paragraphs)
        {
            await Assert.That(p.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("ListParagraph");
            var num = p.ParagraphProperties.NumberingProperties!;
            await Assert.That(num.NumberingLevelReference!.Val?.Value).IsEqualTo(0);
            await Assert.That(num.NumberingId!.Val?.Value).IsNotNull();
        }

        var texts = paragraphs.Select(p => p.GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text).ToList();
        await Assert.That(texts).IsEquivalentTo(["one", "two", "three"]);

        await VerifyDocument(md);
    }

    // Every level used to be emitted at ilvl=0, so nesting was lost: the marker never changed and
    // the indentation the abstractNum's level supplies never applied.
    [Test]
    public async Task NestedListItemsUseTheirNestingLevel()
    {
        var paragraphs = RenderList("- outer\n  - inner\n    - innermost\n- outer two");

        var levels = paragraphs
            .Select(_ => _.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value)
            .ToList();
        await Assert.That(levels).IsEquivalentTo([0, 1, 2, 0]);
    }

    // Word defines nine levels per abstractNum; a deeper list clamps to the last rather than
    // emitting an ilvl with no matching level, which renders unindented and unmarked.
    [Test]
    public async Task NestingDeeperThanNineLevelsClampsToTheLast()
    {
        var markdown = new StringBuilder();
        for (var depth = 0; depth < 12; depth++)
        {
            markdown.Append(new string(' ', depth * 2)).Append("- item").Append('\n');
        }

        var paragraphs = RenderList(markdown.ToString());

        var levels = paragraphs
            .Select(_ => _.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value)
            .ToList();
        await Assert.That(levels).IsEquivalentTo([0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 8, 8]);
    }

    [Test]
    public async Task OrderedListProducesNumberingInstance()
    {
        const string md = "1. first\n2. second";
        var paragraphs = RenderList(md);
        await Assert.That(paragraphs.Count).IsEqualTo(2);

        var numId = paragraphs[0].ParagraphProperties!.NumberingProperties!.NumberingId!.Val?.Value;
        await Assert.That(numId).IsNotNull();

        await VerifyDocument(md);
    }

    [Test]
    public async Task LowerAlphaListUsesLowerLetterFormat()
    {
        const string md = "a. first\nb. second";
        var (paragraphs, renderer) = RenderListWithRenderer(md);
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        await AssertNumberFormat(paragraphs, renderer, NumberFormatValues.LowerLetter);

        await VerifyDocument(md);
    }

    [Test]
    public async Task UpperAlphaListUsesUpperLetterFormat()
    {
        const string md = "A. first\nB. second";
        var (paragraphs, renderer) = RenderListWithRenderer(md);
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        await AssertNumberFormat(paragraphs, renderer, NumberFormatValues.UpperLetter);

        await VerifyDocument(md);
    }

    [Test]
    public async Task LowerRomanListUsesLowerRomanFormat()
    {
        const string md = "i. first\nii. second";
        var (paragraphs, renderer) = RenderListWithRenderer(md);
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        await AssertNumberFormat(paragraphs, renderer, NumberFormatValues.LowerRoman);

        await VerifyDocument(md);
    }

    [Test]
    public async Task UpperRomanListUsesUpperRomanFormat()
    {
        const string md = "I. first\nII. second";
        var (paragraphs, renderer) = RenderListWithRenderer(md);
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        await AssertNumberFormat(paragraphs, renderer, NumberFormatValues.UpperRoman);

        await VerifyDocument(md);
    }

    static async Task AssertNumberFormat(
        List<Paragraph> paragraphs,
        OpenXmlMarkdownRenderer renderer,
        NumberFormatValues expected)
    {
        var numId = paragraphs[0].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
        var numbering = renderer.MainPart.NumberingDefinitionsPart!.Numbering!;
        var instance = numbering.Elements<NumberingInstance>().Single(n => n.NumberID?.Value == numId);
        var abstractId = instance.GetFirstChild<AbstractNumId>()!.Val!.Value;
        var abstractNum = numbering.Elements<AbstractNum>().Single(a => a.AbstractNumberId?.Value == abstractId);
        var level = abstractNum.Elements<Level>().First();
        await Assert.That(level.NumberingFormat!.Val!.Value).IsEqualTo(expected);
    }

    static async Task VerifyDocument(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("list", markdown, styleSource);
        using var stream = new MemoryStream();
        await store.Render("list", new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    static List<Paragraph> RenderList(string markdown)
    {
        var (paragraphs, _) = RenderListWithRenderer(markdown);
        return paragraphs;
    }

    static (List<Paragraph> Paragraphs, OpenXmlMarkdownRenderer Renderer) RenderListWithRenderer(string markdown)
    {
        var list = RendererHarness.FirstBlock<ListBlock>(markdown);
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(list);
        return (renderer.Drain().Cast<Paragraph>().ToList(), renderer);
    }
}
