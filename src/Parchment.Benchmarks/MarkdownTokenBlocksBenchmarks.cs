// Isolates MarkdownTokenBlocks.Apply. In the markdown flow a value that converts to Word blocks —
// `| html`, an HtmlToken, an OpenXmlToken — is parked under a marker during the liquid render and
// swapped in once Markdig has parsed, and inside a loop that is one marker per iteration. Two things
// there used to grow with the square of the loop. Each marker's host was looked for with its own
// walk of every paragraph in the body — O(markers × paragraphs), and a loop grows both. And each
// swap inserted ahead of the host and then removed it, both of which have OpenXML walk every
// sibling before the host. Every host is now found in one walk, what a value produces goes in
// after its host, and the hosts are taken out together at the end. `| html` and an HtmlToken member
// park through the same call, so the filter stands in for both here and lets the sample model be
// used as it is.
//
// PlainLoop is the same loop with the value written as text, which parks nothing. Converting html
// costs about as much again as writing text, so watch the Ratio column level off around 2 instead
// of climbing with LoopItems.
[Config(typeof(BenchmarkConfig))]
public class MarkdownTokenBlocksBenchmarks
{
    TemplateStore plainStore = null!;
    TemplateStore htmlStore = null!;
    ReportContext model = null!;

    [Params(10, 100, 1000)]
    public int LoopItems { get; set; }

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
                Actions = Enumerable.Range(1, LoopItems)
                    .Select(_ => new ActionItem
                    {
                        Title = $"Action item {_}",
                        Detail = $"<p>Details for action {_}.</p>"
                    })
                    .ToList(),
                HasRisks = false
            }
        };

        plainStore = new();
        using var plainStyleSource = new MemoryStream(BuildStyleSource());
        plainStore.RegisterMarkdownTemplate<ReportContext>(plainSource, plainStyleSource);

        htmlStore = new();
        using var htmlStyleSource = new MemoryStream(BuildStyleSource());
        htmlStore.RegisterMarkdownTemplate<ReportContext>(htmlSource, htmlStyleSource);
    }

    [Benchmark(Baseline = true)]
    public async Task PlainLoop()
    {
        using var output = new MemoryStream();
        await plainStore.Render(model, output);
    }

    [Benchmark]
    public async Task HtmlLoop()
    {
        using var output = new MemoryStream();
        await htmlStore.Render(model, output);
    }

    const string plainSource =
        """
        {% for item in Report.Actions %}

        {{ item.Title }}

        {{ item.Detail }}

        {% endfor %}
        """;

    const string htmlSource =
        """
        {% for item in Report.Actions %}

        {{ item.Title }}

        {{ item.Detail | html }}

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
