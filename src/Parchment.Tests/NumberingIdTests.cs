

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A list's <c>w:numId</c> and its definition's <c>w:abstractNumId</c> have to be unique in the
/// numbering part. Markdown lists are numbered by Parchment and html lists by OpenXmlHtml, into the
/// same part, so each has to step past what the other has written since it last looked.
/// </summary>
public partial class NumberingIdTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        [Html]
        public required string HtmlOrdered { get; init; }

        [Html]
        public required string HtmlBullets { get; init; }

        [Markdown]
        public required string MarkdownBullets { get; init; }

        [Markdown]
        public required string MarkdownOrdered { get; init; }

        public required IReadOnlyList<string> Tags { get; init; }

        public required IReadOnlyList<Fragment> Fragments { get; init; }

        public required TokenValue Steps { get; init; }
    }

    public class Fragment
    {
        public required string Html { get; init; }
        public required string Markdown { get; init; }
    }

    [ParchmentBindable]
    public partial class EditableDoc
    {
        [Html]
        [EditableField]
        public required string Body { get; set; }

        [Markdown]
        public required string MarkdownBullets { get; init; }

        [Markdown]
        public required string MarkdownOrdered { get; init; }
    }

    static Doc Model() =>
        new()
        {
            HtmlOrdered = "<ol><li>html one</li><li>html two</li><li>html three</li></ol>",
            HtmlBullets = "<ul><li>html bullet</li></ul>",
            MarkdownBullets = "- md bullet",
            MarkdownOrdered = "1. md one\n2. md two\n3. md three",
            Tags = ["tag"],
            Fragments =
            [
                new()
                {
                    Html = "<ol><li>html 1</li></ol>",
                    Markdown = "- md 1"
                },
                new()
                {
                    Html = "<ul><li>html 2</li></ul>",
                    Markdown = "1. md 2"
                },
                new()
                {
                    Html = "<ol type=\"a\"><li>html 3</li></ol>",
                    Markdown = "- md 3"
                },
                new()
                {
                    Html = "<ul><li>html 4</li></ul><ol><li>html 4 again</li></ol>",
                    Markdown = "a. md 4"
                }
            ],
            Steps = TokenValueHelpers.NumberedList(["step"])
        };

    // The markdown bullet list took abstractNumId 1 and numId 1, the html list read those and took
    // 2 and 3, and the markdown ordered list - counting on from its own 1 - took 2 for its
    // definition as well: one id, two definitions.
    [Test]
    public async Task MarkdownFlowMarkdownListFirst()
    {
        using var stream = await RenderMarkdown(
            """
            - md bullet

            <ol><li>html one</li><li>html two</li><li>html three</li></ol>

            1. md one
            2. md two
            3. md three

            <ul><li>html bullet</li></ul>
            """);

        await AssertLists(
            stream,
            """
            md bullet: bullet
            html one: decimal
            html two: decimal
            html three: decimal
            md one: decimal
            md two: decimal
            md three: decimal
            html bullet: bullet
            """);
    }

    [Test]
    public async Task MarkdownFlowHtmlListFirst()
    {
        using var stream = await RenderMarkdown(
            """
            <ol><li>html one</li></ol>

            - md bullet

            <ul><li>html bullet</li></ul>

            1. md one

            <ol type="a"><li>html alpha</li></ol>

            - md bullet two
            """);

        await AssertLists(
            stream,
            """
            html one: decimal
            md bullet: bullet
            html bullet: bullet
            md one: decimal
            html alpha: lowerLetter
            md bullet two: bullet
            """);
    }

    // The two shapes the collision was first seen in: a third markdown list handed the numId of the
    // html list between the first two.
    [Test]
    public async Task MarkdownFlowListsOfOneKindAroundAnHtmlList()
    {
        using var bullets = await RenderMarkdown(
            """
            - first md bullet

            <ol><li>html one</li><li>html two</li></ol>

            - second md bullet

            Paragraph.

            - third md bullet
            """);

        await AssertLists(
            bullets,
            """
            first md bullet: bullet
            html one: decimal
            html two: decimal
            second md bullet: bullet
            third md bullet: bullet
            """);

        using var ordered = await RenderMarkdown(
            """
            1. md one

            <ul><li>html bullet</li></ul>

            Paragraph.

            1. md two

            Paragraph.

            1. md three
            """);

        await AssertLists(
            ordered,
            """
            md one: decimal
            html bullet: bullet
            md two: decimal
            md three: decimal
            """);
    }

    [Test]
    public async Task MarkdownFlowSeveralHtmlFragments()
    {
        using var stream = await RenderMarkdown(
            """
            {% for i in (1..4) %}
            - md bullet {{ i }}

            <ol><li>html number {{ i }}</li></ol>

            1. md number {{ i }}

            <ul><li>html bullet {{ i }}</li></ul>

            {% endfor %}
            """);

        var expected = new StringBuilder();
        for (var i = 1; i <= 4; i++)
        {
            expected.AppendLine($"md bullet {i}: bullet");
            expected.AppendLine($"html number {i}: decimal");
            expected.AppendLine($"md number {i}: decimal");
            expected.AppendLine($"html bullet {i}: bullet");
        }

        await AssertLists(stream, expected.ToString());
    }

    // An html list inside a list item is converted part way through the markdown list around it.
    // Not validated, because the indent an html block is given inside a list item is written out
    // of schema order, which is a separate matter from the ids.
    [Test]
    public async Task MarkdownFlowHtmlListInsideAListItem()
    {
        using var stream = await RenderMarkdown(
            """
            - md outer

              <ol><li>html in item</li></ol>

              1. md nested

            - md outer two

            <ul><li>html bullet</li></ul>

            a. md alpha
            """);

        await AssertLists(
            stream,
            """
            md outer: bullet
            html in item: decimal
            md nested: decimal
            md outer two: bullet
            html bullet: bullet
            md alpha: lowerLetter
            """,
            validate: false);
    }

    // A parked value is converted after the markdown around it has been numbered, in the order the
    // tokens were written, so an html value and a list token alternate the same way.
    [Test]
    public async Task MarkdownFlowParkedHtmlAndListTokens()
    {
        using var stream = await RenderMarkdown(
            """
            - md bullet

            {{ HtmlBullets }}

            {{ Steps }}

            {{ HtmlOrdered }}

            {{ Steps }}
            """);

        await AssertLists(
            stream,
            """
            md bullet: bullet
            html bullet: bullet
            step: decimal
            html one: decimal
            html two: decimal
            html three: decimal
            step: decimal
            """);
    }

    [Test]
    public async Task DocxFlowMarkdownListFirst()
    {
        using var stream = await RenderDocx(
            """
            {{ MarkdownBullets }}

            {{ HtmlOrdered }}

            {{ MarkdownOrdered }}

            {{ HtmlBullets }}
            """);

        await AssertLists(
            stream,
            """
            md bullet: bullet
            html one: decimal
            html two: decimal
            html three: decimal
            md one: decimal
            md two: decimal
            md three: decimal
            html bullet: bullet
            """);
    }

    // The html lists share a numbering session here, which carries the converter's next id from one
    // conversion to the next and reads the part again only when it ends somewhere new.
    [Test]
    public async Task DocxFlowHtmlListFirst()
    {
        using var stream = await RenderDocx(
            """
            {{ HtmlOrdered }}

            {{ MarkdownBullets }}

            {{ HtmlBullets }}

            {{ MarkdownOrdered }}

            {{ HtmlBullets }}

            {{ MarkdownBullets }}
            """);

        await AssertLists(
            stream,
            """
            html one: decimal
            html two: decimal
            html three: decimal
            md bullet: bullet
            html bullet: bullet
            md one: decimal
            md two: decimal
            md three: decimal
            html bullet: bullet
            md bullet: bullet
            """);
    }

    [Test]
    public async Task DocxFlowSeveralHtmlFragmentsInALoop()
    {
        using var stream = await RenderDocx(
            """
            {% for fragment in Fragments %}

            {{ fragment.Html | html }}

            {{ fragment.Markdown | markdown }}

            {% endfor %}
            """);

        await AssertLists(
            stream,
            """
            html 1: decimal
            md 1: bullet
            html 2: bullet
            md 2: decimal
            html 3: lowerLetter
            md 3: bullet
            html 4: bullet
            html 4 again: decimal
            md 4: lowerLetter
            """);
    }

    // The lists that never pass through markdown: a string collection, and the two list filters.
    [Test]
    public async Task DocxFlowStringListsAndListFilters()
    {
        using var stream = await RenderDocx(
            """
            {{ Tags }}

            {{ HtmlOrdered }}

            {{ Tags | numbered_list }}

            {{ HtmlBullets }}

            {{ Tags | bullet_list }}

            {{ Steps }}
            """);

        await AssertLists(
            stream,
            """
            tag: bullet
            html one: decimal
            html two: decimal
            html three: decimal
            tag: decimal
            html bullet: bullet
            tag: bullet
            step: decimal
            """);
    }

    [Test]
    public async Task DocxFlowInTableCells()
    {
        using var template = DocxTemplateBuilder.BuildWithTable(
            "{{ HtmlOrdered }}",
            "{{ MarkdownOrdered }}",
            bodyParagraphText: "{{ MarkdownBullets }}");
        using var stream = await RenderDocx(template);

        await AssertLists(
            stream,
            """
            md bullet: bullet
            html one: decimal
            html two: decimal
            html three: decimal
            md one: decimal
            md two: decimal
            md three: decimal
            """);
    }

    // One numbering part serves every part of the document, so a header's lists are numbered from
    // the same counters as the body's.
    [Test]
    public async Task DocxFlowAcrossParts()
    {
        using var template = BuildDocx(
            [
                Paragraph("{{ MarkdownBullets }}"),
                Paragraph("{{ HtmlOrdered }}")
            ],
            [
                Paragraph("{{ MarkdownOrdered }}"),
                Paragraph("{{ HtmlBullets }}"),
                Paragraph("{{ MarkdownBullets }}")
            ]);
        using var stream = await RenderDocx(template);

        await AssertLists(
            stream,
            """
            md bullet: bullet
            html one: decimal
            html two: decimal
            html three: decimal
            md one: decimal
            md two: decimal
            md three: decimal
            html bullet: bullet
            md bullet: bullet
            """);
    }

    // A template that arrives with lists of its own, in the shape Word for Mac saves them: the
    // part ends with something other than a w:num. Not validated, because where a definition is
    // written among the template's own is a separate matter from which id it is given.
    [Test]
    public async Task DocxFlowTemplateWithItsOwnNumbering()
    {
        using var template = BuildDocx(
            [
                TemplateListItem("template item"),
                Paragraph("{{ HtmlBullets }}"),
                Paragraph("{{ MarkdownBullets }}"),
                Paragraph("{{ HtmlBullets }}"),
                Paragraph("{{ MarkdownOrdered }}"),
                Paragraph("{{ HtmlOrdered }}"),
                Paragraph("{{ MarkdownBullets }}")
            ],
            numbering: true);
        using var stream = await RenderDocx(template);

        await AssertLists(
            stream,
            """
            template item: upperRoman
            html bullet: bullet
            md bullet: bullet
            html bullet: bullet
            md one: decimal
            md two: decimal
            md three: decimal
            html one: decimal
            html two: decimal
            html three: decimal
            md bullet: bullet
            """,
            validate: false);
    }

    // The other order in such a template, after html that holds no list: OpenXmlHtml has by then
    // counted the template's ids, and has to count the two markdown lists as well before it numbers
    // the html one.
    [Test]
    public async Task DocxFlowTemplateWithItsOwnNumberingMarkdownListsFirst()
    {
        using var template = BuildDocx(
            [
                TemplateListItem("template item"),
                Paragraph("{{ '<p>html paragraph</p>' | html }}"),
                Paragraph("{{ MarkdownBullets }}"),
                Paragraph("{{ MarkdownOrdered }}"),
                Paragraph("{{ HtmlBullets }}"),
                Paragraph("{{ MarkdownBullets }}")
            ],
            numbering: true);
        using var stream = await RenderDocx(template);

        await AssertLists(
            stream,
            """
            template item: upperRoman
            md bullet: bullet
            md one: decimal
            md two: decimal
            md three: decimal
            html bullet: bullet
            md bullet: bullet
            """,
            validate: false);
    }

    // The session OpenXmlHtml is handed carries its next id from one conversion to the next, and
    // it reads the part again only when the part no longer ends with the element the session left
    // there. That says where the last definition was written, not whether one was, so a session is
    // not kept across the ids taken for a markdown list. The trailing element is put back last by
    // hand here, which is where it stays when an instance is written ahead of it.
    [Test]
    public async Task HtmlSessionCountsAgainAfterMarkdownLists()
    {
        using var stream = BuildDocx([], numbering: true);
        using var doc = WordprocessingDocument.Open(stream, true);
        var mainPart = doc.MainDocumentPart!;
        var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;
        var state = new WordNumberingState(mainPart);

        state.ConvertHtml(
            "<p>html paragraph</p>",
            ImagePolicies.AllowAll.BuildSettings(numberingSession: state.GetHtmlSession()));
        state.CreateBulletNumbering();
        state.CreateOrderedNumbering(NumberFormatValues.Decimal);

        var cleanup = numbering.GetFirstChild<NumberingIdMacAtCleanup>()!;
        cleanup.Remove();
        numbering.Append(cleanup);

        state.ConvertHtml(
            "<ul><li>html bullet</li></ul>",
            ImagePolicies.AllowAll.BuildSettings(numberingSession: state.GetHtmlSession()));

        var definitions = numbering
            .Elements<AbstractNum>()
            .Select(_ => _.AbstractNumberId!.Value)
            .ToList();
        var instances = numbering
            .Elements<NumberingInstance>()
            .Select(_ => _.NumberID!.Value)
            .ToList();
        await Assert.That(definitions.Distinct().Count()).IsEqualTo(definitions.Count);
        await Assert.That(instances.Distinct().Count()).IsEqualTo(instances.Count);
    }

    [Test]
    public async Task DocxFlowEditableHtml()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ MarkdownBullets }}

            {{ Body }}

            {{ MarkdownOrdered }}

            {{ MarkdownBullets }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableDoc>(template);

        using var stream = await Render(
            store,
            new EditableDoc
            {
                Body = "<ol><li>html one</li></ol><ul><li>html bullet</li></ul>",
                MarkdownBullets = "- md bullet",
                MarkdownOrdered = "1. md one"
            });

        await AssertLists(
            stream,
            """
            md bullet: bullet
            html one: decimal
            html bullet: bullet
            md one: decimal
            md bullet: bullet
            """);
    }

    [Test]
    public async Task IdsAreTheSameOnEveryRender()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ MarkdownBullets }}

            {{ HtmlOrdered }}

            {{ MarkdownOrdered }}

            {{ HtmlBullets }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var first = await Render(store, Model());
        using var second = await Render(store, Model());

        await Assert.That(second.ToArray()).IsEquivalentTo(first.ToArray());
    }

    // What Word draws beside each list paragraph: the format of the level its numId leads to. That
    // is only an answer when each id names one thing, which is asserted first.
    static async Task AssertLists(MemoryStream stream, string expected, bool validate = true)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart!;
        var numbering = mainPart.NumberingDefinitionsPart!.Numbering!;

        var definitions = numbering
            .Elements<AbstractNum>()
            .Select(_ => _.AbstractNumberId!.Value)
            .ToList();
        var instances = numbering
            .Elements<NumberingInstance>()
            .Select(_ => _.NumberID!.Value)
            .ToList();
        await Assert.That(definitions.Distinct().Count()).IsEqualTo(definitions.Count);
        await Assert.That(instances.Distinct().Count()).IsEqualTo(instances.Count);

        var lists = new List<string>();
        foreach (var paragraph in Roots(mainPart).SelectMany(_ => _.Descendants<Paragraph>()))
        {
            var properties = paragraph.ParagraphProperties?.NumberingProperties;
            if (properties?.NumberingId?.Val?.Value is not { } numId)
            {
                continue;
            }

            var ilvl = properties.NumberingLevelReference?.Val?.Value ?? 0;
            var instance = numbering
                .Elements<NumberingInstance>()
                .Single(_ => _.NumberID!.Value == numId);
            var definition = numbering
                .Elements<AbstractNum>()
                .Single(_ => _.AbstractNumberId!.Value == instance.AbstractNumId!.Val!.Value);
            var level = definition
                .Elements<Level>()
                .Single(_ => _.LevelIndex!.Value == ilvl);
            lists.Add($"{paragraph.InnerText}: {level.NumberingFormat!.Val!.InnerText}");
        }

        await Assert.That(string.Join('\n', lists)).IsEqualTo(expected.ReplaceLineEndings("\n").TrimEnd());

        if (!validate)
        {
            return;
        }

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static IEnumerable<OpenXmlElement> Roots(MainDocumentPart mainPart)
    {
        yield return mainPart.Document!.Body!;
        foreach (var header in mainPart.HeaderParts)
        {
            yield return header.Header!;
        }
    }

    static async Task<MemoryStream> RenderMarkdown(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return await Render(store, Model());
    }

    static async Task<MemoryStream> RenderDocx(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        return await RenderDocx(template);
    }

    static Task<MemoryStream> RenderDocx(MemoryStream template)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return Render(store, Model());
    }

    static async Task<MemoryStream> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildDocx(Paragraph[] body, Paragraph[]? header = null, bool numbering = false)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var documentBody = mainPart.Document.Body!;
            documentBody.Append(body);

            var section = new SectionProperties();
            if (header != null)
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
                headerPart.Header = new(header);
                section.Append(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    });
            }

            if (numbering)
            {
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>("rIdTemplateNumbering");
                numberingPart.Numbering = new(
                    new AbstractNum(
                        new Level(
                            new NumberingFormat
                            {
                                Val = NumberFormatValues.UpperRoman
                            },
                            new LevelText
                            {
                                Val = "%1."
                            })
                        {
                            LevelIndex = 0
                        })
                    {
                        AbstractNumberId = 4
                    },
                    new NumberingInstance(
                        new AbstractNumId
                        {
                            Val = 4
                        })
                    {
                        NumberID = 7
                    },
                    new NumberingIdMacAtCleanup
                    {
                        Val = 7
                    });
            }

            section.Append(
                new PageSize
                {
                    Width = 6500,
                    Height = 8000
                });
            documentBody.Append(section);
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph Paragraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    static Paragraph TemplateListItem(string text) =>
        new(
            new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference
                    {
                        Val = 0
                    },
                    new NumberingId
                    {
                        Val = 7
                    })),
            new Run(new Text(text)));
}
