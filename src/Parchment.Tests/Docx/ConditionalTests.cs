public class ConditionalTests
{
    [Test]
    public async Task IfTrue()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Start

            {% if Customer.IsPreferred %}

            Preferred customer: {{ Customer.Name }}

            {% else %}

            Regular customer: {{ Customer.Name }}

            {% endif %}

            End
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>("conditional", template);
        using var stream = new MemoryStream();
        await store.Render("conditional", SampleData.Invoice(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    public class FlagModel
    {
        public required bool Flag { get; init; }
        public required string Label { get; init; }
    }

    [Test]
    public async Task ElseBranchRenders()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Start

            {% if Flag %}

            Affirmative: {{ Label }}

            {% else %}

            Negative: {{ Label }}

            {% endif %}

            End
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<FlagModel>("else-branch", template);
        using var stream = new MemoryStream();
        await store.Render(
            "else-branch",
            new FlagModel
            {
                Flag = false,
                Label = "fallback"
            },
            stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    public class TwoFlagModel
    {
        public required bool Flag { get; init; }
        public required bool Second { get; init; }
        public required string Label { get; init; }
    }

    [Test]
    public async Task StaticParagraphInChosenBranchIsKept()
    {
        // Static (token-free) paragraphs get no anchors and no scope-tree nodes, so branch
        // content must be kept positionally — an anchor-derived keep-set silently dropped them.
        using var template = DocxTemplateBuilder.Build(
            """
            Before

            {% if Flag %}

            Static kept text.

            {% endif %}

            After
            """);

        var texts = await RenderToTexts(
            template,
            new FlagModel
            {
                Flag = true,
                Label = "unused"
            });
        await Assert.That(texts).IsEquivalentTo(["Before", "Static kept text.", "After"]);
    }

    [Test]
    public async Task StaticOnlyElseBranchIsKept()
    {
        // A static-only else branch has an empty scope-tree body; the else tag anchor alone must
        // be enough to keep its physical range.
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Flag %}

            Affirmative: {{ Label }}

            {% else %}

            Fallback static.

            {% endif %}
            """);

        var texts = await RenderToTexts(
            template,
            new FlagModel
            {
                Flag = false,
                Label = "unused"
            });
        await Assert.That(texts).IsEquivalentTo(["Fallback static."]);
    }

    [Test]
    public async Task MixedStaticAndTokenBranches()
    {
        var content =
            """
            {% if Flag %}

            Intro static.

            Value: {{ Label }}

            Outro static.

            {% else %}

            Else static.

            Else value: {{ Label }}

            {% endif %}
            """;

        using (var template = DocxTemplateBuilder.Build(content))
        {
            var texts = await RenderToTexts(
                template,
                new FlagModel
                {
                    Flag = true,
                    Label = "yes"
                });
            await Assert.That(texts).IsEquivalentTo(["Intro static.", "Value: yes", "Outro static."]);
        }

        using (var template = DocxTemplateBuilder.Build(content))
        {
            var texts = await RenderToTexts(
                template,
                new FlagModel
                {
                    Flag = false,
                    Label = "no"
                });
            await Assert.That(texts).IsEquivalentTo(["Else static.", "Else value: no"]);
        }
    }

    [Test]
    public async Task ElsifChosenKeepsStaticAndRemovesSiblingBranches()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Flag %}

            First: {{ Label }}

            {% elsif Second %}

            Elsif static.

            {% else %}

            Else: {{ Label }}

            {% endif %}
            """);

        var texts = await RenderToTexts(
            template,
            new TwoFlagModel
            {
                Flag = false,
                Second = true,
                Label = "unused"
            });
        await Assert.That(texts).IsEquivalentTo(["Elsif static."]);
    }

    [Test]
    public async Task NestedIfStaticParagraphs()
    {
        var content =
            """
            {% if Flag %}

            Outer before.

            {% if Second %}

            Inner static.

            {% endif %}

            Outer after.

            {% endif %}
            """;

        using (var template = DocxTemplateBuilder.Build(content))
        {
            var texts = await RenderToTexts(
                template,
                new TwoFlagModel
                {
                    Flag = true,
                    Second = true,
                    Label = "unused"
                });
            await Assert.That(texts).IsEquivalentTo(["Outer before.", "Inner static.", "Outer after."]);
        }

        using (var template = DocxTemplateBuilder.Build(content))
        {
            var texts = await RenderToTexts(
                template,
                new TwoFlagModel
                {
                    Flag = true,
                    Second = false,
                    Label = "unused"
                });
            await Assert.That(texts).IsEquivalentTo(["Outer before.", "Outer after."]);
        }
    }

    [Test]
    public async Task StaticParagraphInEliminatedBranchIsRemoved()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Before

            {% if Flag %}

            Eliminated static.

            {% endif %}

            After
            """);

        var texts = await RenderToTexts(
            template,
            new FlagModel
            {
                Flag = false,
                Label = "unused"
            });
        await Assert.That(texts).IsEquivalentTo(["Before", "After"]);
    }

    [Test]
    public async Task TableInsideChosenBranchIsKeptAndSubstituted()
    {
        // Tables are captured positionally like static paragraphs (they have no anchor of their
        // own), and tokens inside table cells must still substitute.
        using var template = BuildTemplateWithTableInIf();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<FlagModel>("if-table-kept", template);

        using var stream = new MemoryStream();
        await store.Render(
            "if-table-kept",
            new FlagModel
            {
                Flag = true,
                Label = "cell value"
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var table = body.Elements<Table>().SingleOrDefault();
        await Assert.That(table).IsNotNull();
        await Assert.That(table!.InnerText).IsEqualTo("Cell: cell value");
    }

    [Test]
    public async Task TableInsideEliminatedBranchIsRemoved()
    {
        using var template = BuildTemplateWithTableInIf();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<FlagModel>("if-table-removed", template);

        using var stream = new MemoryStream();
        await store.Render(
            "if-table-removed",
            new FlagModel
            {
                Flag = false,
                Label = "unused"
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.Elements<Table>().Any()).IsFalse();
    }

    [Test]
    public async Task DuplicateElseIsRejected()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Flag %}

            A

            {% else %}

            B

            {% else %}

            C

            {% endif %}
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<FlagModel>("double-else", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("after '{% else %}'");
    }

    [Test]
    public async Task ElsifAfterElseIsRejected()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Flag %}

            A

            {% else %}

            B

            {% elsif Second %}

            C

            {% endif %}
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<TwoFlagModel>("elsif-after-else", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("after '{% else %}'");
    }

    static async Task<List<string>> RenderToTexts<TModel>(
        Stream template,
        TModel model,
        [CallerMemberName] string name = "")
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<TModel>(name, template);
        using var stream = new MemoryStream();
        await store.Render(name, model!, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Select(_ => string.Concat(_.Descendants<Text>().Select(t => t.Text)))
            .Where(_ => _.Length > 0)
            .ToList();
    }

    static MemoryStream BuildTemplateWithTableInIf()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(
                new Body(
                    BuildParagraph("{% if Flag %}"),
                    new Table(
                        new TableRow(
                            new TableCell(
                                BuildParagraph("Cell: {{ Label }}")))),
                    BuildParagraph("{% endif %}"),
                    new SectionProperties(
                        new PageSize
                        {
                            Width = 6500,
                            Height = 8000
                        })));
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    [Test]
    public async Task NoMatchingBranch_NoElse_AllBranchParagraphsRemoved()
    {
        // Condition is false and there is no else branch. Every paragraph between {% if %} and
        // {% endif %} (plus the open/close anchor paragraphs) must be removed; surrounding
        // "Before"/"After" paragraphs stay.
        using var template = DocxTemplateBuilder.Build(
            """
            Before

            {% if Flag %}

            Affirmative: {{ Label }}

            {% endif %}

            After
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<FlagModel>("no-match-no-else", template);
        using var stream = new MemoryStream();
        await store.Render(
            "no-match-no-else",
            new FlagModel
            {
                Flag = false,
                Label = "ignored"
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var texts = doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Select(_ => string.Concat(_.Descendants<Text>().Select(t => t.Text)))
            .Where(_ => _.Length > 0)
            .ToList();
        await Assert.That(texts).IsEquivalentTo(["Before", "After"]);
    }
}
