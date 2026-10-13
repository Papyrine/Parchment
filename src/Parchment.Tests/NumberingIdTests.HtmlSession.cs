// Html met on the markdown side is converted with the render's numbering session, which carries
// OpenXmlHtml's next id from one conversion to the next instead of having it read the numbering
// part every time. These hold the ids to what a read of the part gives.
public partial class NumberingIdTests
{
    const string pixel = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

    const int rounds = 30;

    static string InterleavedMarkdown()
    {
        var builder = new StringBuilder();
        for (var i = 1; i <= rounds; i++)
        {
            builder.Append(
                $"""
                 - md bullet {i}

                 <ol><li>html number {i}</li></ol>

                 <ul><li>html bullet {i}</li></ul>

                 <ul><li>html again {i}</li></ul>

                 1. md number {i}

                 ![alt]({pixel})

                 <div>plain {i}</div>

                 <ol type="a"><li>html alpha {i}</li></ol>

                 - md second {i}

                 Paragraph {i}.


                 """);
        }

        return builder.ToString();
    }

    // The ids are worked out here the way they are when the part is read for every list: an html
    // list takes the two ids past the highest of either kind, one for a definition of its own and
    // one for its instance; a markdown list takes the instance id past the highest instance, and
    // the definition id past the highest definition the first time its format is used.
    [Test]
    public async Task MarkdownFlowManyInterleavedLists()
    {
        using var stream = await RenderMarkdown(InterleavedMarkdown());

        var highestDefinition = 0;
        var highestInstance = 0;
        int? bulletDefinition = null;
        int? decimalDefinition = null;
        var expected = new StringBuilder();

        void Html(string text)
        {
            var definition = Math.Max(highestDefinition, highestInstance) + 1;
            highestDefinition = definition;
            highestInstance = definition + 1;
            expected.Append($"{text}: num {highestInstance} of {definition}\n");
        }

        void Markdown(string text, ref int? definition)
        {
            definition ??= ++highestDefinition;
            highestInstance++;
            expected.Append($"{text}: num {highestInstance} of {definition}\n");
        }

        for (var i = 1; i <= rounds; i++)
        {
            Markdown($"md bullet {i}", ref bulletDefinition);
            Html($"html number {i}");
            Html($"html bullet {i}");
            Html($"html again {i}");
            Markdown($"md number {i}", ref decimalDefinition);
            Html($"html alpha {i}");
            Markdown($"md second {i}", ref bulletDefinition);
        }

        await Assert.That(ListIds(stream)).IsEqualTo(expected.ToString());
        await AssertDistinctAndValid(stream);
    }

    // Html blocks alone: nothing is numbered by Parchment between them, so one session lasts the
    // whole walk. Each ul still writes a bullet definition of its own, as it did when the part was
    // read for every block.
    [Test]
    public async Task MarkdownFlowManyHtmlBlocks()
    {
        var markdown = new StringBuilder();
        var expected = new StringBuilder();
        for (var i = 0; i < 200; i++)
        {
            var tag = "ol";
            if (i % 2 == 0)
            {
                tag = "ul";
            }

            markdown.Append($"<{tag}><li>html {i}</li></{tag}>\n\n<p>plain {i}</p>\n\n");
            expected.Append($"html {i}: num {2 * i + 2} of {2 * i + 1}\n");
        }

        using var stream = await RenderMarkdown(markdown.ToString());

        await Assert.That(ListIds(stream)).IsEqualTo(expected.ToString());
        await AssertDistinctAndValid(stream);
    }

    // A style source whose numbering part already holds definitions, and ends with the element
    // Word for Mac writes after the instances.
    [Test]
    public async Task MarkdownFlowManyHtmlBlocksIntoANumberedStyleSource()
    {
        using var styleSource = BuildDocx([], numbering: true);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(
            """
            {% for i in (1..20) %}
            <ol><li>html number {{ i }}</li></ol>

            ![alt](data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==)

            <ul><li>html bullet {{ i }}</li></ul>

            - md bullet {{ i }}

            {% endfor %}
            """,
            styleSource);
        using var stream = await Render(store, Model());

        var expected = new StringBuilder();
        for (var i = 1; i <= 20; i++)
        {
            expected.AppendLine($"html number {i}: decimal");
            expected.AppendLine($"html bullet {i}: bullet");
            expected.AppendLine($"md bullet {i}: bullet");
        }

        await AssertLists(stream, expected.ToString(), validate: false);
    }

    // Token values in a markdown template are converted after the walk, by the same session.
    [Test]
    public async Task MarkdownFlowHtmlTokensAmongLists()
    {
        using var stream = await RenderMarkdown(
            """
            <ul><li>block bullet</li></ul>

            {{ HtmlBullets }}

            - md bullet

            {{ HtmlOrdered }}

            {{ Steps }}

            {{ HtmlBullets }}

            <ol><li>block number</li></ol>
            """);

        await AssertLists(
            stream,
            """
            block bullet: bullet
            html bullet: bullet
            md bullet: bullet
            html one: decimal
            html two: decimal
            html three: decimal
            step: decimal
            html bullet: bullet
            block number: decimal
            """);
    }

    // A token is handed the main part and can write definitions of its own anywhere in the
    // numbering part. A session only notices what is written at the end of it, so one is not kept
    // across a token: here the token's instance goes in ahead of the html list's, and the html
    // that follows would otherwise take the id the token took.
    [Test]
    public async Task MarkdownFlowTokenWritingItsOwnDefinitions()
    {
        var model = Model();
        var steps = new OpenXmlToken(context =>
        {
            var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering!;
            var id = numbering
                .Elements<AbstractNum>()
                .Select(_ => _.AbstractNumberId!.Value)
                .Concat(numbering.Elements<NumberingInstance>().Select(_ => _.NumberID!.Value))
                .Max() + 1;
            var first = numbering.GetFirstChild<NumberingInstance>()!;
            numbering.InsertBefore(
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
                    AbstractNumberId = id
                },
                first);
            numbering.InsertBefore(
                new NumberingInstance(
                    new AbstractNumId
                    {
                        Val = id
                    })
                {
                    NumberID = id
                },
                first);
            return
            [
                new Paragraph(
                    new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference
                            {
                                Val = 0
                            },
                            new NumberingId
                            {
                                Val = id
                            })),
                    new Run(new Text("token item")))
            ];
        });

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(
            """
            {{ HtmlOrdered }}

            {{ Steps }}

            {{ HtmlBullets }}

            {{ HtmlOrdered }}
            """,
            styleSource);
        using var stream = await Render(
            store,
            new Doc
            {
                HtmlOrdered = "<ol><li>html one</li></ol>",
                HtmlBullets = model.HtmlBullets,
                MarkdownBullets = model.MarkdownBullets,
                MarkdownOrdered = model.MarkdownOrdered,
                Tags = model.Tags,
                Fragments = model.Fragments,
                Steps = steps
            });

        await AssertLists(
            stream,
            """
            html one: decimal
            token item: upperRoman
            html bullet: bullet
            html one: decimal
            """);
    }

    // In the docx flow the html tokens share one bullet definition through the session, and an
    // html block inside a markdown value writes its own. Converting the block with the same
    // session leaves both as they were.
    [Test]
    public async Task DocxFlowHtmlBlockInMarkdownKeepsItsOwnBulletDefinition()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ HtmlBullets }}

            {{ MarkdownBullets }}

            {{ HtmlBullets }}

            {{ MarkdownOrdered }}

            {{ HtmlBullets }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        var model = Model();
        using var stream = await Render(
            store,
            new Doc
            {
                HtmlOrdered = model.HtmlOrdered,
                HtmlBullets = "<ul><li>token bullet</li></ul>",
                MarkdownBullets = "<ul><li>block one</li></ul>\n\n![alt](" + pixel + ")\n\n<ul><li>block two</li></ul>",
                MarkdownOrdered = "<ul><li>block three</li></ul>\n\n- md bullet",
                Tags = model.Tags,
                Fragments = model.Fragments,
                Steps = model.Steps
            });

        await Assert.That(ListIds(stream))
            .IsEqualTo(
                """
                token bullet: num 2 of 1
                block one: num 4 of 3
                block two: num 6 of 5
                token bullet: num 7 of 1
                block three: num 9 of 8
                md bullet: num 10 of 9
                token bullet: num 11 of 1

                """.ReplaceLineEndings("\n"));
        await AssertDistinctAndValid(stream);
    }

    // The session belongs to one render. A registered template rendered from many threads at once
    // gives every one of them the bytes a render on its own gives.
    [Test]
    public async Task MarkdownFlowConcurrentRendersMatch()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(
            InterleavedMarkdown() + "{{ HtmlBullets }}\n\n{{ Steps }}\n\n{{ HtmlOrdered }}\n",
            styleSource);
        using var alone = await Render(store, Model());
        var expected = alone.ToArray();

        var renders = Enumerable
            .Range(0, 16)
            .Select(_ => Task.Run(
                async () =>
                {
                    using var stream = await Render(store, Model());
                    return stream.ToArray();
                }))
            .ToList();

        foreach (var render in renders)
        {
            var bytes = await render;
            await Assert.That(bytes.AsSpan().SequenceEqual(expected)).IsTrue();
        }
    }

    // Each list paragraph's text with the instance it names and the definition that leads to.
    static string ListIds(MemoryStream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart!;
        var definitions = mainPart
            .NumberingDefinitionsPart!.Numbering!
            .Elements<NumberingInstance>()
            .ToLookup(_ => _.NumberID!.Value, _ => _.AbstractNumId!.Val!.Value);
        var lists = new StringBuilder();
        foreach (var paragraph in mainPart.Document!.Body!.Descendants<Paragraph>())
        {
            if (paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value is { } numId)
            {
                lists.Append($"{paragraph.InnerText}: num {numId} of {string.Join(',', definitions[numId])}\n");
            }
        }

        return lists.ToString();
    }

    static async Task AssertDistinctAndValid(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
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

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }
}
