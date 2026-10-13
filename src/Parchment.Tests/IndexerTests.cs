// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A token that indexes into a member reads an element of it, not the member. It renders as an
/// ordinary value in both flows, whatever the member alone would have been swapped for, and
/// registration follows the path into the element type.
/// </summary>
public partial class IndexerTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        public required int Index { get; init; }
        public required string Key { get; init; }
        public required string[] Tags { get; init; }
        public required List<Line> Lines { get; init; }
        public required List<List<Line>> Groups { get; init; }
        public required Dictionary<string, Line> Lookup { get; init; }

        [ExcelsiorTable]
        public required List<Row> Rows { get; init; }

        [Html]
        public required string Body { get; init; }

        [Markdown]
        public required string Notes { get; init; }

        public required Order Order { get; init; }
        public required List<Order> Orders { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
        public List<string> Tags { get; init; } = [];
    }

    public class Row
    {
        [Column(Order = 1)]
        public required string Item { get; init; }

        [Column(Order = 2)]
        public required int Qty { get; init; }
    }

    public class Order
    {
        public required List<string> Tags { get; init; }

        [ExcelsiorTable]
        public required List<Row> Rows { get; init; }
    }

    public class Budget
    {
        [EditableField]
        public string? Year { get; set; }

        public List<string> Notes { get; set; } = [];
    }

    [ParchmentBindable]
    public partial class Form
    {
        [EditableField]
        public required string Reference { get; set; }

        [EditableField]
        public required List<Budget> Budgets { get; set; }
    }

    static Doc BuildModel() =>
        new()
        {
            Index = 1,
            Key = "spare",
            Tags = ["red", "green"],
            Lines =
            [
                new()
                {
                    Name = "one",
                    Tags = ["a", "b"]
                },
                new()
                {
                    Name = "two"
                }
            ],
            Groups =
            [
                [
                    new()
                    {
                        Name = "first"
                    },
                    new()
                    {
                        Name = "second"
                    }
                ],
                [
                    new()
                    {
                        Name = "third"
                    }
                ]
            ],
            Lookup = new()
            {
                ["spare"] = new()
                {
                    Name = "looked up",
                    Tags = ["x", "yz"]
                }
            },
            Rows =
            [
                new()
                {
                    Item = "widget",
                    Qty = 4
                },
                new()
                {
                    Item = "gadget",
                    Qty = 5
                }
            ],
            Body = "<b>bold</b>",
            Notes = "**strong**",
            Order = new()
            {
                Tags = ["rush", "fragile"],
                Rows =
                [
                    new()
                    {
                        Item = "crate",
                        Qty = 1
                    }
                ]
            },
            Orders =
            [
                new()
                {
                    Tags = ["north", "cold"],
                    Rows =
                    [
                        new()
                        {
                            Item = "ice",
                            Qty = 2
                        }
                    ]
                },
                new()
                {
                    Tags = ["south"],
                    Rows =
                    [
                        new()
                        {
                            Item = "sand",
                            Qty = 3
                        }
                    ]
                }
            ]
        };

    // Alone in its paragraph is where the member itself is swapped for a bullet list or a table,
    // so each token is rendered there as well as beside other text.
    [Test]
    [Arguments("{{ Tags[0] }}", "red")]
    [Arguments("{{ Tags[1] }}", "green")]
    [Arguments("{{ Tags[Index] }}", "green")]
    [Arguments("{{ Tags[0].size }}", "3")]
    [Arguments("{{ Order.Tags[0] }}", "rush")]
    [Arguments("{{ Rows[0].Item }}", "widget")]
    [Arguments("{{ Rows[Index].Qty }}", "5")]
    [Arguments("{{ Order.Rows[0].Item }}", "crate")]
    [Arguments("{{ Lines[0].Name }}", "one")]
    [Arguments("{{ Lines[0].Tags[1] }}", "b")]
    [Arguments("{{ Groups[0][1].Name }}", "second")]
    [Arguments("{{ Lookup[Key].Name }}", "looked up")]
    public async Task IndexedTokenRendersTheElement(string token, string expected)
    {
        await Assert.That(await RenderDocx(token)).IsEqualTo(expected);
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown(token)).IsEqualTo(expected);
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    // Past the end there is no element. Indexing a string has no answer in liquid either, so an
    // [Html] or [Markdown] member indexed that way is nothing rather than the whole formatted value.
    [Test]
    [Arguments("{{ Tags[5] }}")]
    [Arguments("{{ Rows[5].Item }}")]
    [Arguments("{{ Lines[5].Name }}")]
    [Arguments("{{ Body[0] }}")]
    [Arguments("{{ Body[Index] }}")]
    [Arguments("{{ Notes[0] }}")]
    public async Task IndexWithNoElementRendersNothing(string token)
    {
        await Assert.That(await RenderDocx(token)).IsEqualTo("");
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo("[]");
        await Assert.That(await RenderMarkdown(token)).IsEqualTo("");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo("[]");
    }

    [Test]
    public async Task IndexedTokenProducesNoListOrTable()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Tags[0] }}

            {{ Rows[0].Item }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.Descendants<Table>()).IsEmpty();
        await Assert.That(body.Descendants<NumberingProperties>()).IsEmpty();
        await Assert.That(body.Elements<Paragraph>().Select(_ => _.InnerText)).IsEquivalentTo(["red", "widget"]);
    }

    [Test]
    public async Task VariableIndexInsideALoop()
    {
        var docx = await RenderDocx(
            """
            {% for i in (0..1) %}

            {{ Tags[i] }}

            {{ Rows[i].Item }};

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("redwidget;greengadget;");

        var markdown = await RenderMarkdown("{% for i in (0..1) %}{{ Tags[i] }}{{ Rows[i].Item }};{% endfor %}");
        await Assert.That(markdown).IsEqualTo("redwidget;greengadget;");
    }

    [Test]
    public async Task IndexOnALoopVariableMember()
    {
        var docx = await RenderDocx(
            """
            {% for order in Orders %}

            {{ order.Tags[0] }}

            {{ order.Rows[0].Item }};

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("northice;southsand;");

        var markdown = await RenderMarkdown("{% for order in Orders %}{{ order.Tags[0] }}{{ order.Rows[0].Item }};{% endfor %}");
        await Assert.That(markdown).IsEqualTo("northice;southsand;");
    }

    // The loop variable is typed from the element the index reaches, not from the member indexed.
    [Test]
    public async Task IndexedLoopSource()
    {
        var docx = await RenderDocx(
            """
            {% for line in Groups[0] %}

            {{ line.Name }};

            {% endfor %}

            {% for tag in Lines[0].Tags %}

            {{ tag.size }}{{ tag }};

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("first;second;1a;1b;");

        var markdown = await RenderMarkdown(
            "{% for line in Groups[0] %}{{ line.Name }};{% endfor %}{% for tag in Lines[0].Tags %}{{ tag.size }}{{ tag }};{% endfor %}");
        await Assert.That(markdown).IsEqualTo("first;second;1a;1b;");
    }

    // A dictionary key is only known at render, so nothing past it can be checked: the loop
    // variable is accepted untyped instead of being typed as the dictionary's own element.
    [Test]
    public async Task LoopSourceBehindADictionaryKeyIsUntyped()
    {
        var docx = await RenderDocx(
            """
            {% for tag in Lookup[Key].Tags %}

            {{ tag.size }}{{ tag }};

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("1x;2yz;");

        var markdown = await RenderMarkdown("{% for tag in Lookup[Key].Tags %}{{ tag.size }}{{ tag }};{% endfor %}");
        await Assert.That(markdown).IsEqualTo("1x;2yz;");
    }

    [Test]
    [Arguments("{{ Rows[0].Missing }}")]
    [Arguments("{{ Rows[Index].Missing }}")]
    [Arguments("{{ Order.Rows[0].Missing }}")]
    [Arguments("{{ Lines[0].Missing }}")]
    [Arguments("{{ Lines[0].Tags[0].Missing }}")]
    [Arguments("{{ Groups[0][0].Missing }}")]
    [Arguments("{{ Tags[0].Missing }}")]
    public async Task UnknownMemberPastAnIndexFailsRegistration(string token)
    {
        using var template = DocxTemplateBuilder.Build(token);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("'Missing'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(token, styleSource));
        await Assert.That(markdown.Message).Contains("'Missing'");
    }

    [Test]
    public async Task UnknownMemberPastAnIndexInsideALoopFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for order in Orders %}

            {{ order.Rows[0].Missing }}

            {% endfor %}
            """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("'Missing'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(
                "{% for order in Orders %}{{ order.Rows[0].Missing }}{% endfor %}",
                styleSource));
        await Assert.That(markdown.Message).Contains("'Missing'");
    }

    [Test]
    public async Task UnknownMemberOnAnIndexedLoopSourceFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for line in Groups[0] %}

            {{ line.Missing }}

            {% endfor %}
            """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("'Missing'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(
                "{% for line in Groups[0] %}{{ line.Missing }}{% endfor %}",
                styleSource));
        await Assert.That(markdown.Message).Contains("'Missing'");
    }

    // Lines[0] is one line, where the member indexed into is the list of them.
    [Test]
    public async Task IndexedLoopSourceThatIsNotEnumerableFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for part in Lines[0] %}

            x

            {% endfor %}
            """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("'Lines[]' does not resolve to an enumerable type");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(
                "{% for part in Lines[0] %}x{% endfor %}",
                styleSource));
        await Assert.That(markdown.Message).Contains("'Lines[]' does not resolve to an enumerable type");
    }

    // A bracketed string names a member, which is the same token as the dotted spelling.
    [Test]
    public async Task StringLiteralIndexerIsStillTheMember() =>
        await Assert.That(await RenderDocx("{{ Order['Tags'] }}")).IsEqualTo("rushfragile");

    // The header runs through the same dispatch as the body.
    [Test]
    public async Task IndexedTokenInAHeader()
    {
        using var template = DocxTemplateBuilder.Build("{{ Tags[1] }}");
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(
                new Paragraph(
                    new Run(
                        new Text("{{ Tags[0] }}"))));
            mainPart.Document!.Body!.GetFirstChild<SectionProperties>()!.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdHeader"
                });
        }

        template.Position = 0;
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;

        using var rendered = WordprocessingDocument.Open(stream, false);
        var main = rendered.MainDocumentPart!;
        await Assert.That(main.HeaderParts.Single().Header!.InnerText).IsEqualTo("red");
        await Assert.That(main.Document!.Body!.InnerText).IsEqualTo("green");
    }

    // An index into an [EditableField] string is not the field: no content control. A loop over
    // something reached through an element of an editable collection is not the collection: no
    // repeating section.
    [Test]
    public async Task IndexedEditableMemberIsNotTheField()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            [{{ Reference[0] }}]

            {% for note in Budgets[0].Notes %}

            {{ note }};

            {% endfor %}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Form>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new Form
            {
                Reference = "PO-1",
                Budgets =
                [
                    new()
                    {
                        Year = "2026",
                        Notes = ["draft", "final"]
                    }
                ]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.Descendants<SdtElement>()).IsEmpty();
        await Assert.That(body.InnerText).IsEqualTo("[]draft;final;");
    }

    static async Task<string> RenderDocx(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return await RenderText(store);
    }

    static async Task<string> RenderMarkdown(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return await RenderText(store);
    }

    static async Task<string> RenderText(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
