[Config(typeof(BenchmarkConfig))]
public class LoopBenchmarks
{
    TemplateStore store = null!;
    Invoice model = null!;

    [Params(10, 100, 1000)]
    public int LoopItems { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        model = new()
        {
            Number = "INV-001",
            IssueDate = new(2026, 1, 1),
            DueDate = new(2026, 2, 1),
            Currency = "USD",
            Customer = new()
            {
                Name = "Test",
                Email = "t@t",
                Address = "1 Main St"
            },
            Lines = Enumerable.Range(1, LoopItems)
                .Select(
                    _ => new LineItem
                {
                    Description = $"Item {_}",
                    Quantity = _,
                    UnitPrice = 10m
                })
                .ToList()
        };

        store = new();
        using var template = BuildTemplate();
        store.RegisterDocxTemplate<Invoice>(template);
    }

    [Benchmark]
    public async Task RenderLoop()
    {
        using var output = new MemoryStream();
        await store.Render(model, output);
    }

    static MemoryStream BuildTemplate()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(
                new Body(
                    Para("{% for line in Lines %}"),
                    Para("{{ line.Description }}: {{ line.Quantity }} x {{ line.UnitPrice }}"),
                    Para("{% endfor %}")));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            styles.Append(
                new Style
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = "Normal",
                        Default = true
                    }
                    .AppendChild(
                        new StyleName
                        {
                            Val = "Normal"
                        })
                    .Parent!);
            stylesPart.Styles = styles;
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph Para(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
