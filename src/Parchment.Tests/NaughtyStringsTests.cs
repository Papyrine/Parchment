public partial class NaughtyStringsTests
{
    [ParchmentBindable]
    public partial class NaughtyModel
    {
        public required string Single { get; init; }
        public required IReadOnlyList<NaughtyItem> Items { get; init; }
    }

    public class NaughtyItem
    {
        public required string Value { get; init; }
    }

    static NaughtyModel BuildModel() =>
        new()
        {
            Single = string.Join(" | ", TheNaughtyStrings.All),
            Items = TheNaughtyStrings.All
                .Select(_ => new NaughtyItem {Value = _})
                .ToList()
        };

    [Test]
    public async Task Docx()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Single: {{ Single }}

            {% for item in Items %}

            {{ item.Value }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<NaughtyModel>(template);

        using var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task Markdown()
    {
        var markdownSource =
            """
            # {{ Single }}

            {% for item in Items %}
            - {{ item.Value }}
            {% endfor %}
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<NaughtyModel>(
            markdownSource,
            styleSource: styleSource);

        using var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
