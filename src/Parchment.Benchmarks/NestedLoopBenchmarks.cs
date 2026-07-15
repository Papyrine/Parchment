// Isolates Fix 2 (per-loop IdentifierVisitor.Collect). The inner `{% for t in Tags %}` loop is
// entered once per outer `Lines` iteration. The old code ran IdentifierVisitor.Collect over the
// loop source AST on every one of those entries (allocating a visitor + path lists) just to probe
// the editable-collection map. The source's dotted reference is now resolved once at tree-build
// time and the probe is skipped entirely when no editable collections exist, so the per-entry
// allocation disappears — visible as reduced Allocated in the MemoryDiagnoser column.
[Config(typeof(BenchmarkConfig))]
public class NestedLoopBenchmarks
{
    TemplateStore store = null!;
    Invoice model = null!;

    [Params(50, 500)]
    public int OuterItems { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        model = new()
        {
            Number = "INV-001",
            IssueDate = new(2026, 1, 1),
            DueDate = new(2026, 2, 1),
            Currency = "USD",
            Tags = ["priority", "net-30", "urgent"],
            Customer = new()
            {
                Name = "Test",
                Email = "t@t",
                Address = "1 Main St"
            },
            Lines = Enumerable.Range(1, OuterItems)
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
        store.RegisterDocxTemplate<Invoice>("nested", template);
    }

    [Benchmark]
    public async Task RenderNestedLoop()
    {
        using var output = new MemoryStream();
        await store.Render("nested", model, output);
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
                    Para("{{ line.Description }}"),
                    Para("{% for t in Tags %}"),
                    Para("{{ t }}"),
                    Para("{% endfor %}"),
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
