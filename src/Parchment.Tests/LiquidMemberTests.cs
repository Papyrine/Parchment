// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// Registration has to agree with the render about which members of a collection or a string
/// bind. Fluid hands a template an array, a dictionary or a string rather than the CLR object, so
/// what binds is liquid's own <c>size</c> / <c>first</c> / <c>last</c> and none of the CLR members.
/// </summary>
public partial class LiquidMemberTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required string Name { get; init; }
        public required List<Line> Lines { get; init; }
        public required string[] Tags { get; init; }
        public required Dictionary<string, string> Lookup { get; init; }
        public required LineBag Bag { get; init; }
        public required LineList Listed { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
        public IReadOnlyList<Line> Children { get; init; } = [];
    }

    // A model's own collection type is still an array to a template, so its members do not bind.
    public class LineBag :
        IEnumerable<Line>
    {
        List<Line> lines = [];

        public string Label { get; init; } = "label";

        public int Count => lines.Count;

        public void Add(Line line) =>
            lines.Add(line);

        public IEnumerator<Line> GetEnumerator() =>
            lines.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
            GetEnumerator();
    }

    // Deriving from List<T> brings the list's own members in through the base chain, the indexer
    // among them.
    public class LineList :
        List<Line>
    {
        public string Label { get; init; } = "label";
    }

    static Doc BuildModel() =>
        new()
        {
            Name = "Alice",
            Lines =
            [
                new()
                {
                    Name = "one",
                    Children =
                    [
                        new()
                        {
                            Name = "child"
                        }
                    ]
                },
                new()
                {
                    Name = "two"
                },
                new()
                {
                    Name = "three"
                }
            ],
            Tags = ["red", "green"],
            Lookup = new()
            {
                ["key"] = "value"
            },
            Bag =
            [
                new()
                {
                    Name = "bagged"
                }
            ],
            Listed =
            [
                new()
                {
                    Name = "listed"
                }
            ]
        };

    [Test]
    [Arguments("{{ Lines.size }}", "3")]
    [Arguments("{{ Lines.first.Name }}", "one")]
    [Arguments("{{ Lines.last.Name }}", "three")]
    [Arguments("{{ Tags.size }}", "2")]
    [Arguments("{{ Tags.first }}", "red")]
    [Arguments("{{ Tags.last }}", "green")]
    [Arguments("{{ Tags.first.size }}", "3")]
    [Arguments("{{ Name.size }}", "5")]
    [Arguments("{{ Name.first }}", "A")]
    [Arguments("{{ Name.last }}", "e")]
    [Arguments("{{ Lines.first.Name.size }}", "3")]
    [Arguments("{{ Lookup.size }}", "1")]
    [Arguments("{{ Bag.size }}", "1")]
    [Arguments("{{ Bag.first.Name }}", "bagged")]
    [Arguments("{{ Listed.size }}", "1")]
    [Arguments("{{ Listed.first.Name }}", "listed")]
    public async Task LiquidMemberRegistersAndRenders(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    [Test]
    public async Task SizeDrivesAConditional()
    {
        var docx = await RenderDocx(
            """
            {% if Lines.size > 2 %}

            many

            {% endif %}
            """);
        await Assert.That(docx).IsEqualTo("many");

        var markdown = await RenderMarkdown("{% if Lines.size > 2 %}many{% endif %}");
        await Assert.That(markdown).IsEqualTo("many");
    }

    [Test]
    public async Task FirstIsALoopSource()
    {
        var docx = await RenderDocx(
            """
            {% for child in Lines.first.Children %}

            {{ child.Name }}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("child");

        var markdown = await RenderMarkdown("{% for child in Lines.first.Children %}{{ child.Name }}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("child");
    }

    [Test]
    public async Task SizeOfALoopVariableMember()
    {
        var markdown = await RenderMarkdown("{% for line in Lines %}{{ line.Name.size }}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("335");
    }

    // Every one of these is a CLR member, or a liquid member in the wrong case, or a liquid member
    // the value does not have. The first test pins that Fluid binds none of them, which is the only
    // reason the second may reject them.
    [Test]
    [Arguments("Lines.Count")]
    [Arguments("Lines.Capacity")]
    [Arguments("Lines.Size")]
    [Arguments("Lines.First")]
    [Arguments("Tags.Length")]
    [Arguments("Tags.Rank")]
    [Arguments("Name.Length")]
    [Arguments("Name.Size")]
    [Arguments("Lookup.Count")]
    [Arguments("Lookup.Keys")]
    [Arguments("Lookup.first")]
    [Arguments("Bag.Count")]
    [Arguments("Bag.Label")]
    [Arguments("Listed.Count")]
    [Arguments("Listed.Label")]
    public async Task ClrMemberRendersBlankInFluid(string path)
    {
        SharedFluid.Parser.TryParse($"[{{{{ {path} }}}}]", out var template, out _);
        var context = new TemplateContext(BuildModel(), SharedFluid.Options, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(context)).IsEqualTo("[]");
    }

    [Test]
    [Arguments("Lines.Count")]
    [Arguments("Lines.Capacity")]
    [Arguments("Lines.Size")]
    [Arguments("Lines.First")]
    [Arguments("Tags.Length")]
    [Arguments("Tags.Rank")]
    [Arguments("Name.Length")]
    [Arguments("Name.Size")]
    [Arguments("Lookup.Count")]
    [Arguments("Lookup.Keys")]
    [Arguments("Lookup.first")]
    [Arguments("Bag.Count")]
    [Arguments("Bag.Label")]
    [Arguments("Listed.Count")]
    [Arguments("Listed.Label")]
    public async Task ClrMemberFailsRegistrationNamingSize(string path)
    {
        var member = path[(path.IndexOf('.') + 1)..];

        using var template = DocxTemplateBuilder.Build($"{{{{ {path} }}}}");
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains($"'{member}'");
        await Assert.That(docx.Message).Contains("'size'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>($"{{{{ {path} }}}}", styleSource));
        await Assert.That(markdown.Message).Contains($"'{member}'");
        await Assert.That(markdown.Message).Contains("'size'");
    }

    [Test]
    public async Task CountInAConditionFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Lines.Count > 0 %}

            some

            {% endif %}
            """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("'size'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>("{% if Lines.Count > 0 %}some{% endif %}", styleSource));
        await Assert.That(markdown.Message).Contains("'size'");
    }

    // `first` is typed as the element, so a typo past it is still caught.
    [Test]
    public async Task UnknownMemberThroughFirstFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build("{{ Lines.first.Nope }}");
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("Nope");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>("{{ Lines.last.Nope }}", styleSource));
        await Assert.That(markdown.Message).Contains("Nope");
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
