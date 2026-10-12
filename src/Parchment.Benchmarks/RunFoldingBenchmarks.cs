// Isolates the run folding in OpenXmlMarkdownRenderer.FlushParagraph. MarkdownEncoder escapes every
// punctuation character in a bound value and Markdig ends a literal at each escape, so a value
// like "Item-1 (rev. 1.2)" reaches the renderer as a run per fragment, and each adjacent pair is
// asked whether its run properties match before the two fold into one. That question used to be
// answered by serialising both w:rPr to strings — a writer and a string per side, per pair — and
// inside emphasis every fragment carries an rPr, so one bold value serialised the same w:b a
// dozen times over. RunPropertiesComparer walks the two trees instead.
//
// Table is a pipe table with the value in bold in one cell of each row. Prose is a paragraph per
// item with the value inside several kinds of emphasis. Watch Allocated: the strings were most of
// what a render of either allocated.
[Config(typeof(BenchmarkConfig))]
public class RunFoldingBenchmarks
{
    TemplateStore tableStore = null!;
    TemplateStore proseStore = null!;
    ReportContext model = null!;

    [Params(20, 200)]
    public int Items { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        model = new()
        {
            Report = new()
            {
                Title = "Quarterly Review",
                Author = "Alex Chen",
                Date = new(2026, 7, 1),
                Summary = "Platform reliability trended upward this quarter.",
                Findings = [],
                Actions = Enumerable.Range(1, Items)
                    .Select(_ => new ActionItem
                    {
                        Title = $"Item-{_} (rev. 1.2)",
                        Detail = $"Re-test [{_}]: on-call hand-over, 99.9% up-time!"
                    })
                    .ToList(),
                HasRisks = false
            }
        };

        tableStore = new();
        using var tableStyleSource = new MemoryStream(BuildStyleSource());
        tableStore.RegisterMarkdownTemplate<ReportContext>(tableSource, tableStyleSource);

        proseStore = new();
        using var proseStyleSource = new MemoryStream(BuildStyleSource());
        proseStore.RegisterMarkdownTemplate<ReportContext>(proseSource, proseStyleSource);
    }

    [Benchmark]
    public async Task Table()
    {
        using var output = new MemoryStream();
        await tableStore.Render(model, output);
    }

    [Benchmark]
    public async Task Prose()
    {
        using var output = new MemoryStream();
        await proseStore.Render(model, output);
    }

    const string tableSource =
        """
        | Item | Detail |
        | --- | --- |
        {% for item in Report.Actions -%}
        | **{{ item.Title }}** | {{ item.Detail }} |
        {% endfor %}
        """;

    const string proseSource =
        """
        {% for item in Report.Actions %}

        **{{ item.Title }}** needs _{{ item.Detail }}_ before ~~{{ item.Title }}~~ and **_{{ item.Detail }}_** close.

        {% endfor %}
        """;

    static byte[] BuildStyleSource()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(new Paragraph()));

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

        return stream.ToArray();
    }
}
