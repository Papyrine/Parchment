// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A string assigned to a <see cref="TokenValue"/> member is plain text. It has to render as the
/// same string does from a <c>string</c> member, wherever the token sits and in either flow.
/// </summary>
public partial class TokenValueTextTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public string? Plain { get; init; } = "plain text";
        public TokenValue Rich { get; init; } = "plain text";
        public IReadOnlyList<TokenValue> Items { get; init; } = [];
        public IReadOnlyList<Line> Lines { get; init; } = [];
    }

    public class Line
    {
        public required TokenValue Note { get; init; }
    }

    // AddFilter mutates process-wide static state, so the name is unique to this class and is
    // registered once rather than once per test.
    static TokenValueTextTests() =>
        TemplateStore.AddFilter("token_value_text_test_wrap", Wrap);

    static ValueTask<FluidValue> Wrap(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        TokenValue token = input.ToStringValue();
        return new(new ObjectValue(token));
    }

    [Test]
    [Arguments("{{ Rich }}", "plain text")]
    [Arguments("Value: {{ Rich }} end", "Value: plain text end")]
    [Arguments("{{ Rich }}, {{ Plain }}, {{ Rich }}", "plain text, plain text, plain text")]
    public async Task RendersTheText(string content, string expected)
    {
        await Assert.That((await RenderDocx(content, new())).InnerText).IsEqualTo(expected);
        await Assert.That((await RenderMarkdown(content, new())).InnerText).IsEqualTo(expected);
    }

    // Escaping, sanitising and line breaks are whatever a string member gets, because the value is
    // one by the time it is written. So the two paragraphs are compared as xml rather than as text.
    [Test]
    [Arguments("a <b>&amp;</b> \"quoted\"")]
    [Arguments("line one\nline two\r\nline three")]
    [Arguments("  padded\ttabbed  ")]
    [Arguments("")]
    public async Task WrittenAsTheSameStringIs(string value)
    {
        var body = await RenderDocx(
            """
            {{ Plain }}

            {{ Rich }}

            before {{ Plain }} after

            before {{ Rich }} after
            """,
            new()
            {
                Plain = value,
                Rich = value
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs[1].OuterXml).IsEqualTo(paragraphs[0].OuterXml);
        await Assert.That(paragraphs[3].OuterXml).IsEqualTo(paragraphs[2].OuterXml);
    }

    [Test]
    public async Task NewlineBecomesALineBreak()
    {
        var model = new Doc
        {
            Rich = "one\ntwo"
        };

        foreach (var body in new[]
                 {
                     await RenderDocx("{{ Rich }}", model),
                     await RenderMarkdown("{{ Rich }}", model)
                 })
        {
            var paragraph = body.Elements<Paragraph>().Single();
            await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
            await Assert.That(paragraph.InnerText).IsEqualTo("onetwo");
        }
    }

    // The conversion operator is declared for a string, and a null one is a null string to a
    // template: nil, which writes nothing and which `default` replaces.
    [Test]
    [Arguments("{{ Rich }}", "")]
    [Arguments("[{{ Rich }}]", "[]")]
    [Arguments("[{{ Rich | upcase }}]", "[]")]
    [Arguments("[{{ Rich | default: 'none' }}]", "[none]")]
    public async Task NullTextRendersAsANullStringDoes(string content, string expected)
    {
        var model = new Doc
        {
            Plain = null,
            Rich = (string)null!
        };

        await Assert.That((await RenderDocx(content, model)).InnerText).IsEqualTo(expected);
        await Assert.That((await RenderMarkdown(content, model)).InnerText).IsEqualTo(expected);

        var plain = content.Replace("Rich", "Plain");
        await Assert.That((await RenderDocx(plain, model)).InnerText).IsEqualTo(expected);
        await Assert.That((await RenderMarkdown(plain, model)).InnerText).IsEqualTo(expected);
    }

    // A condition sees the string as well, so the operand compares, and is true or false, exactly
    // as the same value in a string member would: an empty string is true in liquid, nil is not.
    [Test]
    [Arguments("Rich == 'plain text'", "plain text", "yes")]
    [Arguments("Rich == 'other'", "plain text", "no")]
    [Arguments("Rich contains 'plain'", "plain text", "yes")]
    [Arguments("Rich", "plain text", "yes")]
    [Arguments("Rich == ''", "", "yes")]
    [Arguments("Rich", "", "yes")]
    [Arguments("Rich", null, "no")]
    [Arguments("Rich == nil", null, "yes")]
    public async Task ConditionSeesTheText(string condition, string? value, string expected)
    {
        var model = new Doc
        {
            Plain = value,
            Rich = value!
        };

        foreach (var operand in new[]
                 {
                     condition,
                     condition.Replace("Rich", "Plain")
                 })
        {
            var docx = await RenderDocx(
                $$"""
                  {% if {{operand}} %}

                  yes

                  {% else %}

                  no

                  {% endif %}
                  """,
                model);
            await Assert.That(docx.InnerText).IsEqualTo(expected);

            var markdown = await RenderMarkdown($"{{% if {operand} %}}yes{{% else %}}no{{% endif %}}", model);
            await Assert.That(markdown.InnerText).IsEqualTo(expected);
        }
    }

    [Test]
    [Arguments("upcase", "PLAIN TEXT")]
    [Arguments("size", "10")]
    [Arguments("append: '!'", "plain text!")]
    [Arguments("replace: 'plain', 'rich'", "rich text")]
    public async Task FilterReceivesTheText(string filter, string expected)
    {
        var content = $"{{{{ Rich | {filter} }}}}";
        await Assert.That((await RenderDocx(content, new())).InnerText).IsEqualTo(expected);
        await Assert.That((await RenderMarkdown(content, new())).InnerText).IsEqualTo(expected);
    }

    // Text is still text the template can ask to have read as markup, as it can of a string.
    [Test]
    public async Task MarkdownFilterReadsTheTextAsMarkdown()
    {
        var model = new Doc
        {
            Rich = "**bold**"
        };

        foreach (var body in new[]
                 {
                     await RenderDocx("{{ Rich | markdown }}", model),
                     await RenderMarkdown("{{ Rich | markdown }}", model)
                 })
        {
            var run = body.Descendants<Run>().Single();
            await Assert.That(run.InnerText).IsEqualTo("bold");
            await Assert.That(run.RunProperties?.Bold).IsNotNull();
        }
    }

    // A loop item is converted on its way out of the collection, and a member of a loop item on
    // its way out of the accessor, so neither arrives as a token.
    [Test]
    public async Task TextInALoop()
    {
        var model = new Doc
        {
            Items = ["one", "two"],
            Lines =
            [
                new()
                {
                    Note = "first"
                },
                new()
                {
                    Note = "second"
                }
            ]
        };

        var docx = await RenderDocx(
            """
            {% for item in Items %}

            {{ item }}

            {% endfor %}

            {% for line in Lines %}

            Note: {{ line.Note }}

            {% endfor %}

            Joined: {{ Items | join: ', ' }}
            """,
            model);
        await Assert.That(Texts(docx)).IsEqualTo("one|two|Note: first|Note: second|Joined: one, two");

        var markdown = await RenderMarkdown(
            "{% for item in Items %}{{ item }}, {% endfor %}{% for line in Lines %}{{ line.Note }}, {% endfor %}{{ Items | join: '+' }}",
            model);
        await Assert.That(markdown.InnerText.Trim()).IsEqualTo("one, two, first, second, one+two");
    }

    // The member can be text for one item and markup for the next, which is what the base type is
    // for. Each is dispatched by what it holds.
    [Test]
    public async Task TextAndMarkupInTheSameLoop()
    {
        var body = await RenderDocx(
            """
            {% for item in Items %}

            {{ item }}

            {% endfor %}
            """,
            new()
            {
                Items = ["**one**", new MarkdownToken("**two**")]
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(Texts(body)).IsEqualTo("**one**|two");
        await Assert.That(paragraphs[0].Descendants<Bold>().Any()).IsFalse();
        await Assert.That(paragraphs[1].Descendants<Bold>().Any()).IsTrue();
    }

    [Test]
    public async Task TextInATableCell()
    {
        using var template = DocxTemplateBuilder.BuildWithTable("Label: {{ Rich }}", "{{ Rich }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, new());
        using var doc = WordprocessingDocument.Open(stream, false);
        var cells = doc.MainDocumentPart!.Document!.Body!
            .Descendants<TableCell>()
            .Select(_ => _.InnerText);
        await Assert.That(string.Join('|', cells)).IsEqualTo("Label: plain text|plain text");
    }

    [Test]
    public async Task TextInHeaderFooterAndNotes()
    {
        using var template = BuildWithParts("Body: {{ Rich }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, new());
        await AssertParts(stream, "Body: plain text");
    }

    // The markdown flow replaces the body and binds every other part of the style source through
    // the docx flow's runner, so a header there is the docx flow as far as this is concerned.
    [Test]
    public async Task TextInAMarkdownStyleSourceHeaderFooterAndNotes()
    {
        using var styleSource = BuildWithParts("replaced by the markdown");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("Body: {{ Rich }}", styleSource);

        using var stream = await Render(store, new());
        await AssertParts(stream, "Body: plain text");
    }

    // The other kinds were never part of this: each is dispatched by its type whether or not it has
    // the paragraph to itself, and none is left for Fluid to turn into a string.
    [Test]
    public async Task MarkupKindsDispatchInline()
    {
        (TokenValue Value, string Expected)[] kinds =
        [
            (new MarkdownToken("**markdown**"), "a markdown b"),
            (new HtmlToken("<b>html</b>"), "a html b"),
            (new OpenXmlToken(_ => [new Paragraph(new Run(new Text("openxml")))]), "a openxml b"),
            (new MutateToken((paragraph, _) => paragraph.Append(new Run(new Text("mutated")))), "a  bmutated")
        ];

        foreach (var (value, expected) in kinds)
        {
            var body = await RenderDocx(
                "a {{ Rich }} b",
                new()
                {
                    Rich = value
                });
            await Assert.That(body.InnerText).IsEqualTo(expected);
        }
    }

    // A filter can hand back a TokenValue of its own, wrapped in an ObjectValue, and no value
    // converter sees that one. Fluid falls back to ToString for an object it has no other answer
    // for, so that has to be the text and not the type name.
    [Test]
    public async Task TextTokenFromAFilterRendersItsText()
    {
        const string content = "[{{ Plain | token_value_text_test_wrap }}]";
        await Assert.That((await RenderDocx(content, new())).InnerText).IsEqualTo("[plain text]");
        await Assert.That((await RenderMarkdown(content, new())).InnerText).IsEqualTo("[plain text]");
    }

    // Text reaches a template as a string, and Fluid answers `size`, `first` and `last` on one.
    // Markup reaches the docx flow as an object, which answers none of them, and which of the two a
    // member holds is only known at the render. This pins both halves, which is the reason the
    // next test may reject them.
    [Test]
    [Arguments("Rich.size")]
    [Arguments("Rich.first")]
    [Arguments("Rich.last")]
    public async Task MemberOfATokenValueDependsOnWhatItHolds(string path)
    {
        SharedFluid.Parser.TryParse($"[{{{{ {path} }}}}]", out var template, out _);

        var text = new TemplateContext(new Doc(), SharedFluid.Options, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(text)).IsNotEqualTo("[]");

        var markup = new TemplateContext(
            new Doc
            {
                Rich = new MarkdownToken("plain text")
            },
            SharedFluid.Options,
            allowModelMembers: true);
        await Assert.That(await template.RenderAsync(markup)).IsEqualTo("[]");
    }

    // So a TokenValue member stays what its declared type says to every validator: an object with
    // no members, which is also not something to loop over. Accepting the string's members would
    // register a template that renders blank whenever the member holds markup. The generator's
    // half of this is in LiquidMemberGeneratorTests.
    [Test]
    [Arguments("{{ Rich.size }}", "'size'")]
    [Arguments("{{ Rich.first }}", "'first'")]
    [Arguments("{{ Rich.last }}", "'last'")]
    [Arguments("{% for item in Rich %}", "'Rich'")]
    public async Task MemberOfATokenValueFailsRegistration(string token, string named)
    {
        var docxContent = token;
        var markdownContent = token;
        if (token.StartsWith("{% for"))
        {
            docxContent = $"{token}\n\n{{{{ item }}}}\n\n{{% endfor %}}";
            markdownContent = $"{token}{{{{ item }}}}{{% endfor %}}";
        }

        using var template = DocxTemplateBuilder.Build(docxContent);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains(named);

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(markdownContent, styleSource));
        await Assert.That(markdown.Message).Contains(named);
    }

    static async Task<Body> RenderDocx(string content, Doc model)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return await RenderBody(store, model);
    }

    static async Task<Body> RenderMarkdown(string markdown, Doc model)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return await RenderBody(store, model);
    }

    static async Task<Body> RenderBody(TemplateStore store, Doc model)
    {
        using var stream = await Render(store, model);
        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    static async Task<MemoryStream> Render(TemplateStore store, Doc model)
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static string Texts(Body body) =>
        string.Join('|', body.Elements<Paragraph>().Select(_ => _.InnerText));

    static async Task AssertParts(Stream stream, string expectedBody)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(main.Document!.Body!.InnerText).IsEqualTo(expectedBody);
        await Assert.That(main.HeaderParts.Single().Header!.InnerText).IsEqualTo("Header: plain text");
        await Assert.That(main.FooterParts.Single().Footer!.InnerText).IsEqualTo("Footer: plain text");
        await Assert.That(main.FootnotesPart!.Footnotes!.InnerText).IsEqualTo("Footnote: plain text");
        await Assert.That(main.EndnotesPart!.Endnotes!.InnerText).IsEqualTo("Endnote: plain text");
    }

    static MemoryStream BuildWithParts(string bodyText)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(BuildParagraph("Header: {{ Rich }}"));

            var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter");
            footerPart.Footer = new(BuildParagraph("Footer: {{ Rich }}"));

            var footnotesPart = mainPart.AddNewPart<FootnotesPart>("rIdFootnotes");
            footnotesPart.Footnotes = new(
                new Footnote(BuildParagraph("Footnote: {{ Rich }}"))
                {
                    Id = 1
                });

            var endnotesPart = mainPart.AddNewPart<EndnotesPart>("rIdEndnotes");
            endnotesPart.Endnotes = new(
                new Endnote(BuildParagraph("Endnote: {{ Rich }}"))
                {
                    Id = 1
                });

            body.Append(BuildParagraph(bodyText));
            body.Append(
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdFooter"
                    },
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));
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
}
