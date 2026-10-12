// A markdown render that uses none of the features the flow has whole-document passes for: no
// `[TOC]`, no `[SECTION]`, no `{#id}` bookmark, no html and no parked token. Every paragraph is
// still asked whether it is a `[TOC]` or `[SECTION]` marker, and that used to build a string from
// the paragraph's text to compare with the marker - twice per paragraph, table cells included. The
// finished body was then walked element by element twice more, once to renumber bookmarks and once
// to find section markers, when the walk that produced it already knew it held neither. A paragraph
// that does not open with a bracket is now turned away without allocating, and both walks are
// skipped when the renderer wrote nothing they look for.
//
// Prose is the paragraphs alone; WithTable adds a pipe table of the same number of rows, whose
// cells are paragraphs too. Allocated is the column to watch: the marker checks were a third of it
// for prose and a sixth with the table.
[Config(typeof(BenchmarkConfig))]
public class MarkdownMarkerFreeBenchmarks
{
    TemplateStore proseStore = null!;
    TemplateStore tableStore = null!;
    ReportContext model = null!;

    [Params(100, 1000)]
    public int Paragraphs { get; set; }

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
                Actions = Enumerable.Range(1, Paragraphs)
                    .Select(_ => new ActionItem
                    {
                        Title = $"Action item {_}",
                        Detail = $"Details for action {_}"
                    })
                    .ToList(),
                HasRisks = false
            }
        };

        proseStore = new();
        using var proseStyleSource = new MemoryStream(BuildStyleSource());
        proseStore.RegisterMarkdownTemplate<ReportContext>(proseSource, proseStyleSource);

        tableStore = new();
        using var tableStyleSource = new MemoryStream(BuildStyleSource());
        tableStore.RegisterMarkdownTemplate<ReportContext>(tableSource, tableStyleSource);
    }

    [Benchmark(Baseline = true)]
    public async Task Prose()
    {
        using var output = new MemoryStream();
        await proseStore.Render(model, output);
    }

    [Benchmark]
    public async Task WithTable()
    {
        using var output = new MemoryStream();
        await tableStore.Render(model, output);
    }

    const string proseSource =
        """
        # {{ Report.Title }}

        {% for item in Report.Actions %}

        {{ item.Title }} was raised in the review and is tracked to completion. {{ item.Detail }}.

        {% endfor %}
        """;

    const string tableSource =
        """
        # {{ Report.Title }}

        {% for item in Report.Actions %}

        {{ item.Title }} was raised in the review and is tracked to completion. {{ item.Detail }}.

        {% endfor %}

        | Action | Detail |
        | --- | --- |
        {% for item in Report.Actions -%}
        | {{ item.Title }} | {{ item.Detail }} |
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
