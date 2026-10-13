

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// <c>w:numbering</c> is a strict sequence: the <c>w:numPicBullet</c> elements, then the
/// <c>w:abstractNum</c> elements, then the <c>w:num</c> elements, then <c>w:numIdMacAtCleanup</c>.
/// A list rendered into a numbering part that already holds picture bullets or the cleanup element
/// has to put its definitions inside that sequence. Word refuses to open a document whose numbering
/// part has a <c>w:abstractNum</c> ahead of a <c>w:numPicBullet</c>.
/// </summary>
public partial class NumberingSequenceTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        public required IReadOnlyList<string> Tags { get; init; }

        [Markdown]
        public required string Notes { get; init; }
    }

    // A bullet list, an ordered list that starts at 3 and so carries a start override on its own
    // instance, and an ordered list of another format, which takes an abstract definition of its
    // own.
    static Doc Model() =>
        new()
        {
            Tags = ["one", "two"],
            Notes =
                """
                - a
                - b

                3. c
                4. d

                text

                i. e
                ii. f
                """
        };

    static readonly string[] sequence = ["numPicBullet", "abstractNum", "num", "numIdMacAtCleanup"];

    [Test]
    [MatrixDataSource]
    public async Task DocxFlowKeepsTheSequence(
        [Matrix(true, false)] bool pictureBullet,
        [Matrix(true, false)] bool definitions,
        [Matrix(true, false)] bool cleanup)
    {
        using var template = WithNumbering(
            DocxTemplateBuilder.Build(
                """
                {{ Tags }}

                {{ Tags | numbered_list }}

                {{ Notes }}
                """),
            pictureBullet,
            definitions,
            cleanup);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        await AssertSequence(store, pictureBullet, definitions, cleanup);
    }

    [Test]
    [MatrixDataSource]
    public async Task MarkdownFlowKeepsTheSequence(
        [Matrix(true, false)] bool pictureBullet,
        [Matrix(true, false)] bool definitions,
        [Matrix(true, false)] bool cleanup)
    {
        using var styleSource = WithNumbering(DocxTemplateBuilder.Build(), pictureBullet, definitions, cleanup);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(
            """
            {{ Tags | bullet_list }}

            {{ Tags | numbered_list }}

            {{ Notes }}
            """,
            styleSource);

        await AssertSequence(store, pictureBullet, definitions, cleanup);
    }

    static async Task AssertSequence(TemplateStore store, bool pictureBullet, bool definitions, bool cleanup)
    {
        using var stream = new MemoryStream();
        await store.Render(Model(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        var names = numbering.ChildElements.Select(_ => _.LocalName).ToList();
        var ordered = names.OrderBy(_ => Array.IndexOf(sequence, _));
        await Assert.That(string.Join(' ', names)).IsEqualTo(string.Join(' ', ordered));

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();

        // What the template brought is still there.
        await Assert.That(names.Count(_ => _ == "numPicBullet")).IsEqualTo(Count(pictureBullet));
        await Assert.That(names.Count(_ => _ == "numIdMacAtCleanup")).IsEqualTo(Count(cleanup));
        if (definitions)
        {
            var own = numbering.Elements<NumberingInstance>().Single(_ => _.NumberID?.Value == 1);
            await Assert.That(own.AbstractNumId?.Val?.Value).IsEqualTo(0);
        }

        // Every list paragraph resolves to one instance, and that to one abstract definition.
        var items = doc.MainDocumentPart.Document!.Body!
            .Descendants<Paragraph>()
            .Where(_ => _.ParagraphProperties?.NumberingProperties != null)
            .ToList();
        await Assert.That(items.Count).IsEqualTo(10);
        foreach (var item in items)
        {
            var abstractId = Instance(numbering, item).AbstractNumId!.Val!.Value;
            await Assert.That(numbering.Elements<AbstractNum>().Count(_ => _.AbstractNumberId?.Value == abstractId)).IsEqualTo(1);
        }

        // The start override is on the instance of the list that starts at 3, wherever that instance
        // had to go.
        var third = Instance(numbering, items.Single(_ => _.InnerText == "c"));
        await Assert.That(third.GetFirstChild<LevelOverride>()?.StartOverrideNumberingValue?.Val?.Value).IsEqualTo(3);
    }

    static int Count(bool present)
    {
        if (present)
        {
            return 1;
        }

        return 0;
    }

    static NumberingInstance Instance(Numbering numbering, Paragraph item)
    {
        var numId = item.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
        return numbering.Elements<NumberingInstance>().Single(_ => _.NumberID?.Value == numId);
    }

    static MemoryStream WithNumbering(MemoryStream template, bool pictureBullet, bool definitions, bool cleanup)
    {
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var numbering = new Numbering();
            if (pictureBullet)
            {
                numbering.Append(
                    new NumberingPictureBullet(new PictureBulletBase())
                    {
                        NumberingPictureBulletId = 0
                    });
            }

            if (definitions)
            {
                numbering.Append(
                    new AbstractNum(
                        new Level(
                            new StartNumberingValue
                            {
                                Val = 1
                            },
                            new NumberingFormat
                            {
                                Val = NumberFormatValues.Decimal
                            },
                            new LevelText
                            {
                                Val = "%1."
                            })
                        {
                            LevelIndex = 0
                        })
                    {
                        AbstractNumberId = 0
                    },
                    new NumberingInstance(
                        new AbstractNumId
                        {
                            Val = 0
                        })
                    {
                        NumberID = 1
                    });
            }

            if (cleanup)
            {
                numbering.Append(
                    new NumberingIdMacAtCleanup
                    {
                        Val = 1
                    });
            }

            doc.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>("rIdTemplateNumbering").Numbering = numbering;
        }

        template.Position = 0;
        return template;
    }
}
