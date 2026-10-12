// ReSharper disable PartialTypeWithSinglePart
public partial class ListBlockRendererTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

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
            markdown.Append(' ', depth * 2).Append("- item").Append('\n');
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

    // The start number was ignored: "5. five" rendered as 1. A list starting at the format's first
    // value says so as well, or it would carry on from the list before it.
    [Test]
    [Arguments("1. first\n2. second", 1)]
    [Arguments("5. five\n6. six", 5)]
    [Arguments("0. zero\n1. one", 0)]
    [Arguments("a. first\nb. second", 1)]
    [Arguments("c. third\nd. fourth", 3)]
    [Arguments("C. third\nD. fourth", 3)]
    [Arguments("i. first\nii. second", 1)]
    [Arguments("iv. fourth\nv. fifth", 4)]
    [Arguments("IV. fourth\nV. fifth", 4)]
    public async Task OrderedListStartsAtTheAuthoredNumber(string md, int expected)
    {
        var (paragraphs, renderer) = RenderListWithRenderer(md);

        await Assert.That(ReadStartOverride(paragraphs[0], renderer)).IsEqualTo(expected);
    }

    [Test]
    public async Task NestedOrderedListStartsAtItsOwnNumber()
    {
        var (paragraphs, renderer) = RenderListWithRenderer("3. outer\n\n   7. inner\n   8. inner two\n4. outer two");

        var starts = paragraphs
            .Select(_ => ReadStartOverride(_, renderer))
            .ToList();
        await Assert.That(starts).IsEquivalentTo(new int?[] { 3, 7, 7, 3 });

        // The override sits on the level the list's items use, and nowhere else.
        var numbering = renderer.MainPart.NumberingDefinitionsPart!.Numbering!;
        var overrides = numbering
            .Descendants<LevelOverride>()
            .Select(_ => _.LevelIndex!.Value)
            .ToList();
        await Assert.That(overrides).IsEquivalentTo([0, 1]);
    }

    [Test]
    public async Task OrderedListStartRendersSchemaValidNumbering()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>("5. five\n6. six\n\ntext\n\n1. one\n2. two", styleSource);
        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var starts = numbering
            .Descendants<StartOverrideNumberingValue>()
            .Select(_ => _.Val!.Value)
            .ToList();
        await Assert.That(starts).IsEquivalentTo([5, 1]);

        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator
            .Validate(document)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    // Every list has its own w:num, but one abstractNum is shared per format, and Word carries a
    // single count across the instances of an abstractNum unless an instance restarts it. Three
    // lists came out as 1,2,3 / 4,5 / 6,7.
    [Test]
    public async Task SeparateOrderedListsEachRestart()
    {
        using var stream = await Render("1. a\n2. b\n3. c\n\ntext\n\n1. d\n2. e\n\ntext\n\n1. f\n2. g");
        using var document = WordprocessingDocument.Open(stream, false);

        await Assert.That(ListRestarts(document)).IsEqualTo("0:1 0:1 0:1");
        await AssertSchemaValid(document);
    }

    // A nested list is an instance of its own as well, so the one under the second parent carried
    // on from the one under the first: a,b then c,d.
    [Test]
    public async Task NestedOrderedListsRestartUnderEachParent()
    {
        const string md =
            """
            1. top one
               a. inner one-a
               b. inner one-b
            2. top two
               a. inner two-a
               b. inner two-b

            Break paragraph.

            1. again one
               1. nested decimal one
               2. nested decimal two
            2. again two
               1. nested decimal one (second parent)
            """;
        using var stream = await Render(md);
        using var document = WordprocessingDocument.Open(stream, false);

        await Assert.That(ListRestarts(document)).IsEqualTo("0:1 1:1 1:1 0:1 1:1 1:1");

        // One override per list, on the level its items use and nowhere else.
        var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var overrideCounts = numbering
            .Elements<NumberingInstance>()
            .Select(_ => _.Elements<LevelOverride>().Count())
            .ToList();
        await Assert.That(overrideCounts).IsEquivalentTo([1, 1, 1, 1, 1, 1]);
        await AssertSchemaValid(document);
    }

    // A bullet has nothing to count, so its instance stays bare.
    [Test]
    public async Task BulletListsCarryNoRestart()
    {
        using var stream = await Render("- a\n- b\n\ntext\n\n- c\n  1. d\n- e");
        using var document = WordprocessingDocument.Open(stream, false);

        await Assert.That(ListRestarts(document)).IsEqualTo("0: 0: 1:1");
    }

    // Every paragraph of an item used to carry the numPr, so a continuation paragraph took a marker
    // of its own and an ordered list counted it as the next item.
    [Test]
    public async Task ContinuationParagraphIsNotNumbered()
    {
        var (paragraphs, renderer) = RenderListWithRenderer("1. first\n\n   continuation\n\n2. second");

        await Assert.That(paragraphs.Count).IsEqualTo(3);
        await Assert.That(paragraphs[0].ParagraphProperties!.NumberingProperties).IsNotNull();
        await Assert.That(paragraphs[2].ParagraphProperties!.NumberingProperties).IsNotNull();

        var continuation = paragraphs[1].ParagraphProperties!;
        await Assert.That(continuation.NumberingProperties).IsNull();
        await Assert.That(continuation.ParagraphStyleId!.Val?.Value).IsEqualTo("ListParagraph");
        await Assert.That(continuation.Indentation!.Left?.Value).IsEqualTo(LevelIndent(paragraphs[0], renderer).Left?.Value);
        await Assert.That(continuation.Indentation.Hanging).IsNull();
    }

    [Test]
    public async Task NestedContinuationParagraphAlignsWithItsOwnLevel()
    {
        var (paragraphs, renderer) = RenderListWithRenderer("- outer\n  - inner\n\n    inner continuation\n\n  outer continuation");

        await Assert.That(paragraphs.Count).IsEqualTo(4);
        await Assert.That(paragraphs[2].ParagraphProperties!.NumberingProperties).IsNull();
        await Assert.That(paragraphs[3].ParagraphProperties!.NumberingProperties).IsNull();
        await Assert.That(paragraphs[2].ParagraphProperties!.Indentation!.Left?.Value)
            .IsEqualTo(LevelIndent(paragraphs[1], renderer).Left?.Value);
        await Assert.That(paragraphs[3].ParagraphProperties!.Indentation!.Left?.Value)
            .IsEqualTo(LevelIndent(paragraphs[0], renderer).Left?.Value);
        await Assert.That(paragraphs[2].ParagraphProperties!.Indentation!.Left?.Value).IsEqualTo("960");
        await Assert.That(paragraphs[3].ParagraphProperties!.Indentation!.Left?.Value).IsEqualTo("480");
    }

    [Test]
    public async Task LooseListNumbersEveryItem()
    {
        var paragraphs = RenderList("1. first\n\n2. second\n\n3. third");

        await Assert.That(paragraphs.Count).IsEqualTo(3);
        foreach (var paragraph in paragraphs)
        {
            await Assert.That(paragraph.ParagraphProperties!.NumberingProperties).IsNotNull();
            await Assert.That(paragraph.ParagraphProperties.Indentation).IsNull();
        }
    }

    // An item that opens with a code block still has to be counted, so its first paragraph takes
    // the marker even though it is not the first block.
    [Test]
    public async Task ParagraphAfterLeadingCodeBlockIsNumbered()
    {
        var paragraphs = RenderList("1. ```\n   code\n   ```\n\n   text\n\n   more text");

        await Assert.That(paragraphs.Count).IsEqualTo(3);
        await Assert.That(paragraphs[0].ParagraphProperties!.NumberingProperties).IsNull();
        await Assert.That(paragraphs[1].ParagraphProperties!.NumberingProperties).IsNotNull();
        await Assert.That(paragraphs[2].ParagraphProperties!.NumberingProperties).IsNull();
    }

    [Test]
    public async Task ContinuationParagraphKeepsTheListStyle()
    {
        var paragraphs = RenderList("{.ListStyle}\n- first\n\n  continuation");

        await Assert.That(paragraphs.Count).IsEqualTo(2);
        await Assert.That(paragraphs[1].ParagraphProperties!.NumberingProperties).IsNull();
        await Assert.That(paragraphs[1].ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("ListStyle");
    }

    // Both fixes together: the override belongs to the list's instance, so the continuation
    // paragraph neither carries a numPr nor disturbs the start.
    [Test]
    public async Task ContinuationParagraphLeavesTheStartOverrideAlone()
    {
        var (paragraphs, renderer) = RenderListWithRenderer("5. five\n\n   continuation\n\n6. six");

        await Assert.That(paragraphs.Count).IsEqualTo(3);
        await Assert.That(ReadStartOverride(paragraphs[0], renderer)).IsEqualTo(5);
        await Assert.That(paragraphs[1].ParagraphProperties!.NumberingProperties).IsNull();
        await Assert.That(paragraphs[2].ParagraphProperties!.NumberingProperties!.NumberingId!.Val?.Value)
            .IsEqualTo(paragraphs[0].ParagraphProperties!.NumberingProperties!.NumberingId!.Val?.Value);

        var numbering = renderer.MainPart.NumberingDefinitionsPart!.Numbering!;
        await Assert.That(numbering.Descendants<LevelOverride>().Count()).IsEqualTo(1);
    }

    static int? ReadStartOverride(Paragraph paragraph, OpenXmlMarkdownRenderer renderer)
    {
        var properties = paragraph.ParagraphProperties!.NumberingProperties!;
        var numId = properties.NumberingId!.Val!.Value;
        var ilvl = properties.NumberingLevelReference!.Val!.Value;
        var numbering = renderer.MainPart.NumberingDefinitionsPart!.Numbering!;
        var instance = numbering.Elements<NumberingInstance>().Single(_ => _.NumberID?.Value == numId);
        var levelOverride = instance.Elements<LevelOverride>().SingleOrDefault(_ => _.LevelIndex?.Value == ilvl);
        return levelOverride?.StartOverrideNumberingValue?.Val?.Value;
    }

    // One entry per list in document order, as "level:start": the level its items sit at and the
    // number its instance restarts that level at, empty where the instance does not restart it.
    static string ListRestarts(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart!;
        var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;
        var seen = new HashSet<int>();
        var lists = new List<string>();
        foreach (var properties in mainPart.Document!.Body!.Descendants<NumberingProperties>())
        {
            var numId = properties.NumberingId!.Val!.Value;
            if (!seen.Add(numId))
            {
                continue;
            }

            var ilvl = properties.NumberingLevelReference!.Val!.Value;
            var instance = numbering.Elements<NumberingInstance>().Single(_ => _.NumberID?.Value == numId);
            var levelOverride = instance.Elements<LevelOverride>().SingleOrDefault(_ => _.LevelIndex?.Value == ilvl);
            lists.Add($"{ilvl}:{levelOverride?.StartOverrideNumberingValue?.Val?.Value}");
        }

        return string.Join(' ', lists);
    }

    static async Task AssertSchemaValid(WordprocessingDocument document)
    {
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator
            .Validate(document)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static async Task<MemoryStream> Render(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource);
        var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        return stream;
    }

    static Indentation LevelIndent(Paragraph numbered, OpenXmlMarkdownRenderer renderer)
    {
        var properties = numbered.ParagraphProperties!.NumberingProperties!;
        var numId = properties.NumberingId!.Val!.Value;
        var ilvl = properties.NumberingLevelReference!.Val!.Value;
        var numbering = renderer.MainPart.NumberingDefinitionsPart!.Numbering!;
        var instance = numbering.Elements<NumberingInstance>().Single(_ => _.NumberID?.Value == numId);
        var abstractId = instance.GetFirstChild<AbstractNumId>()!.Val!.Value;
        var abstractNum = numbering.Elements<AbstractNum>().Single(_ => _.AbstractNumberId?.Value == abstractId);
        var level = abstractNum.Elements<Level>().Single(_ => _.LevelIndex?.Value == ilvl);
        return level.PreviousParagraphProperties!.Indentation!;
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
        store.RegisterMarkdownTemplate<EmptyModel>(markdown, styleSource);
        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
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
