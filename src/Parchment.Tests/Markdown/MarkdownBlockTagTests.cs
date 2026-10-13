// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A markdown template is ordinary liquid, so <c>unless</c>, <c>case</c>, <c>cycle</c>,
/// <c>tablerow</c> and the rest are all usable. Registration has to walk them the way the render
/// does. Generator counterpart: <c>MarkdownBlockTagGeneratorTests</c>, over the same
/// <see cref="MarkdownBlockTagCases"/>.
/// </summary>
public partial class MarkdownBlockTagTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required string Name { get; init; }
        public bool Flag { get; init; }
        public bool Off { get; init; }
        public int Count { get; init; }
        public required List<Line> Lines { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
        public List<Line> Children { get; init; } = [];
    }

    static Doc BuildModel() =>
        new()
        {
            Name = "Alice",
            Flag = true,
            Count = 2,
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
            ]
        };

    [Test]
    [MethodDataSource(typeof(MarkdownBlockTagCases), nameof(MarkdownBlockTagCases.Accepted))]
    public async Task RegistersAndRenders(string template)
    {
        using var document = await Render(template);
        await Assert.That(document.MainDocumentPart!.Document!.Body).IsNotNull();
    }

    [Test]
    [MethodDataSource(typeof(MarkdownBlockTagCases), nameof(MarkdownBlockTagCases.Rejected))]
    public async Task FailsRegistration(string template, string path)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(template, styleSource));
        await Assert.That(exception.Message).Contains($"[model path: {path}]");
    }

    [Test]
    public async Task VariableIntroducedInsideABranchRenders()
    {
        await Assert.That(await RenderText("{% unless Off %}{% assign n = 1 %}{% endunless %}{{ n }}")).IsEqualTo("1");
        await Assert.That(await RenderText("{% case Count %}{% when 2 %}{% assign n = 1 %}{% endcase %}{{ n }}")).IsEqualTo("1");
        await Assert.That(await RenderText("{% unless Off %}{% capture c %}x{% endcapture %}{% endunless %}{{ c }}")).IsEqualTo("x");
    }

    [Test]
    public async Task CycleRenders() =>
        await Assert.That(await RenderText("{% for l in Lines %}{% cycle Name, l.Name %} {% endfor %}")).IsEqualTo("Alice two Alice");

    // The reason a counter's name is rejected as a token: Fluid keeps it apart from variables.
    [Test]
    public async Task CounterIsNotAVariableInFluid()
    {
        SharedFluid.Parser.TryParse("{% increment n %}{% decrement m %}[{{ n }}{{ m }}]", out var template, out _);
        var context = new TemplateContext(BuildModel(), SharedFluid.Options, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(context)).EndsWith("[]");
    }

    [Test]
    public async Task TableRowInsideATableBecomesAWordTable()
    {
        using var document = await Render(
            """
            before

            <table>
            {% tablerow l in Lines cols: Count %}{{ l.Name }} {{ tablerowloop.index }}{% endtablerow %}
            </table>

            after
            """);
        var body = document.MainDocumentPart!.Document!.Body!;
        var table = body.Elements<Table>().Single();
        var rows = table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>().Select(_ => _.InnerText).ToList())
            .ToList();

        await Assert.That(rows.Count).IsEqualTo(2);
        await Assert.That(rows[0]).IsEquivalentTo(["one 1", "two 2"]);
        // Liquid leaves the last row short; the html conversion squares it off.
        await Assert.That(rows[1]).IsEquivalentTo(["three 3", ""]);
        await Assert.That(body.Elements<Paragraph>().Select(_ => _.InnerText).Where(_ => _.Length > 0))
            .IsEquivalentTo(["before", "after"]);

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .Select(_ => _.Description)
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task TableRowSourceThatIsNotEnumerableFailsRegistration()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>("<table>{% tablerow c in Count %}x{% endtablerow %}</table>", styleSource));
        await Assert.That(exception.Message).Contains("Loop source 'Count' does not resolve to an enumerable type.");
    }

    static async Task<string> RenderText(string markdown)
    {
        using var document = await Render(markdown);
        return document.MainDocumentPart!.Document!.Body!.InnerText;
    }

    static async Task<WordprocessingDocument> Render(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }
}
