

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A <see cref="MutateToken"/> delegate can do anything to its paragraph. The other tokens in that
/// paragraph are located by offsets measured before it ran, so they are all substituted first and
/// the delegate gets the paragraph last.
/// </summary>
public partial class MutateTokenTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required string Name;
        public required string Count;
        public required MutateToken Token;
        public required MutateToken Second;
        public required HtmlToken Block;
        public required List<string> Items;
    }

    const string threeTokens = "{{ Name }} {{ Token }} {{ Count }}";

    static void Prepend(Paragraph paragraph, string text) =>
        paragraph.InsertAfter(TextRun(text), paragraph.ParagraphProperties);

    static Run TextRun(string text) =>
        new(
            new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

    static Doc Model(Action<Paragraph> mutate, Action<Paragraph>? second = null) =>
        new()
        {
            Name = "Alice",
            Count = "3",
            Token = new((paragraph, _) => mutate(paragraph)),
            Second = new((paragraph, _) => second?.Invoke(paragraph)),
            Block = new("<p>one</p><p>two</p>"),
            Items = ["a", "b"]
        };

    // Was "Alice" followed by "Name }}  3": the prefix moved every earlier offset along.
    [Test]
    public async Task InsertBeforeLeavesEarlierTokensWhole()
    {
        var paragraphs = await Render(threeTokens, Model(_ => Prepend(_, "PREFIX ")));

        await Assert.That(paragraphs).IsEquivalentTo(["PREFIX Alice  3"]);
    }

    [Test]
    public async Task InsertAfter()
    {
        var paragraphs = await Render(threeTokens, Model(_ => _.Append(TextRun(" SUFFIX"))));

        await Assert.That(paragraphs).IsEquivalentTo(["Alice  3 SUFFIX"]);
    }

    // The delegate sees the values of the other tokens, not their source text.
    [Test]
    public async Task ChangesEarlierText()
    {
        var paragraphs = await Render(
            threeTokens,
            Model(paragraph =>
            {
                foreach (var text in paragraph.Descendants<Text>())
                {
                    text.Text = text.Text.ToUpperInvariant();
                }
            }));

        await Assert.That(paragraphs).IsEquivalentTo(["ALICE  3"]);
    }

    [Test]
    public async Task WrapsEarlierRuns()
    {
        var paragraphs = await Render(
            threeTokens,
            Model(paragraph =>
            {
                var hyperlink = new Hyperlink
                {
                    Anchor = "somewhere"
                };
                foreach (var run in paragraph.Elements<Run>().ToList())
                {
                    run.Remove();
                    hyperlink.Append(run);
                }

                paragraph.Append(hyperlink);
            }));

        await Assert.That(paragraphs).IsEquivalentTo(["Alice  3"]);
    }

    [Test]
    public async Task ReplacesTheWholeContent()
    {
        var paragraphs = await Render(
            threeTokens,
            Model(paragraph =>
            {
                paragraph.RemoveAllChildren<Run>();
                paragraph.Append(TextRun("replaced"));
            }));

        await Assert.That(paragraphs).IsEquivalentTo(["replaced"]);
    }

    [Test]
    public async Task RemovesTheParagraph()
    {
        var paragraphs = await Render(
            $"""
             {threeTokens}

             {"{{ Name }}"}
             """,
            Model(_ => _.Remove()));

        await Assert.That(paragraphs).IsEquivalentTo(["Alice"]);
    }

    [Test]
    public async Task BookmarkAroundEverything()
    {
        using var stream = await RenderStream(
            threeTokens,
            Model(paragraph =>
            {
                paragraph.InsertAfter(
                    new BookmarkStart
                    {
                        Name = "mine",
                        Id = "900"
                    },
                    paragraph.ParagraphProperties);
                paragraph.Append(
                    new BookmarkEnd
                    {
                        Id = "900"
                    });
            }));
        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("Alice  3");

        // Paired by whatever id the bookmark ends up with, and around both values.
        var start = paragraph.Elements<BookmarkStart>().Single();
        var end = paragraph.Elements<BookmarkEnd>().Single();
        await Assert.That(start.Name?.Value).IsEqualTo("mine");
        await Assert.That(end.Id?.Value).IsEqualTo(start.Id?.Value);
        await Assert.That(start.ElementsBefore().OfType<Run>().Any()).IsFalse();
        await Assert.That(end.ElementsAfter().OfType<Run>().Any()).IsFalse();
    }

    // The delegates run in the order the tokens are reached, the last in the paragraph first.
    [Test]
    public async Task TwoInOneParagraph()
    {
        var paragraphs = await Render(
            "{{ Token }} {{ Name }} {{ Second }} {{ Count }}",
            Model(_ => Prepend(_, "A"), _ => Prepend(_, "B")));

        await Assert.That(paragraphs).IsEquivalentTo(["AB Alice  3"]);
    }

    [Test]
    public async Task InALoopBody()
    {
        var paragraphs = await Render(
            """
            {% for item in Items %}

            {{ item }} {{ Token }} {{ Count }}

            {% endfor %}
            """,
            Model(_ => Prepend(_, "PREFIX ")));

        await Assert.That(paragraphs).IsEquivalentTo(["PREFIX a  3", "PREFIX b  3"]);
    }

    [Test]
    public async Task InAConditionalBranch()
    {
        var paragraphs = await Render(
            $$$"""
               {% if Name == 'Alice' %}

               {{{threeTokens}}}

               {% endif %}
               """,
            Model(_ => Prepend(_, "PREFIX ")));

        await Assert.That(paragraphs).IsEquivalentTo(["PREFIX Alice  3"]);
    }

    [Test]
    public async Task InATableCell()
    {
        using var template = DocxTemplateBuilder.BuildWithTable("label", threeTokens);
        var paragraphs = await Render(template, Model(_ => Prepend(_, "PREFIX ")));

        await Assert.That(paragraphs).IsEquivalentTo(["label", "PREFIX Alice  3"]);
    }

    // A block-shaped sibling splits the paragraph at its own token. The half a mutate token's
    // text ended up in is the paragraph its delegate gets.
    [Test]
    public async Task AfterASplitGetsTheHalfItSitsIn()
    {
        var paragraphs = await Render(
            "{{ Name }} {{ Block }} tail {{ Token }}",
            Model(_ => Prepend(_, "PREFIX")));

        await Assert.That(paragraphs).IsEquivalentTo(["Alice ", "one", "two", "PREFIX tail "]);
    }

    [Test]
    public async Task BeforeASplitGetsTheHalfItSitsIn()
    {
        var paragraphs = await Render(
            "{{ Token }} {{ Name }} {{ Block }} tail",
            Model(_ => Prepend(_, "PREFIX")));

        await Assert.That(paragraphs).IsEquivalentTo(["PREFIX Alice ", "one", "two", " tail"]);
    }

    static async Task<List<string>> Render(string content, Doc model)
    {
        using var template = DocxTemplateBuilder.Build(content);
        return await Render(template, model);
    }

    static async Task<List<string>> Render(Stream template, Doc model)
    {
        using var stream = await RenderStream(template, model);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!
            .Descendants<Paragraph>()
            .Select(_ => _.InnerText)
            .ToList();
    }

    static async Task<MemoryStream> RenderStream(string content, Doc model)
    {
        using var template = DocxTemplateBuilder.Build(content);
        return await RenderStream(template, model);
    }

    static async Task<MemoryStream> RenderStream(Stream template, Doc model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
                .Validate(doc)
                .Select(_ => _.Description)
                .ToList();
            await Assert.That(errors).IsEmpty();
        }

        stream.Position = 0;
        return stream;
    }
}
