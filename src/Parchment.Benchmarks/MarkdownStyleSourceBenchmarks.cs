// Isolates what a markdown render pays for its style source. The markdown replaces the body, but a
// style source authored in Word usually still carries the sample document it was designed around,
// and every render used to parse that body only to discard it, then load the settings part to
// confirm the compatibility setting was there. Both depend on the style source alone, so
// registration now bakes them into the snapshot a render clones: the body is already empty apart
// from the section it preserves.
//
// BodyParagraphs is the size of the sample document left in the style source; the markdown and the
// model are the same throughout. Watch Mean and Allocated stay level as BodyParagraphs grows, where
// they used to climb with it.
[Config(typeof(BenchmarkConfig))]
public class MarkdownStyleSourceBenchmarks
{
    TemplateStore store = null!;
    ReportContext model = null!;

    [Params(0, 100, 1000)]
    public int BodyParagraphs { get; set; }

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
                Actions = Enumerable.Range(1, 20)
                    .Select(_ => new ActionItem
                    {
                        Title = $"Action item {_}",
                        Detail = $"Details for action {_}."
                    })
                    .ToList(),
                HasRisks = false
            }
        };

        store = new();
        using var styleSource = new MemoryStream(BuildStyleSource(BodyParagraphs));
        store.RegisterMarkdownTemplate<ReportContext>(source, styleSource);
    }

    [Benchmark]
    public async Task Render()
    {
        using var output = new MemoryStream();
        await store.Render(model, output);
    }

    const string source =
        """
        # {{ Report.Title }}

        {{ Report.Summary }}

        | Action | Detail |
        | --- | --- |
        {% for item in Report.Actions -%}
        | {{ item.Title }} | {{ item.Detail }} |
        {% endfor %}
        """;

    static byte[] BuildStyleSource(int bodyParagraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            for (var i = 0; i < bodyParagraphs; i++)
            {
                body.Append(
                    new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines
                            {
                                After = "120"
                            }),
                        new Run(
                            new RunProperties(new Bold()),
                            new Text($"Sample heading {i}. ")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }),
                        new Run(
                            new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor."))));
            }

            body.Append(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 11906,
                        Height = 16838
                    }));
            mainPart.Document = new(body);

            // As Word writes it: a settings part of its own, already declaring the mode.
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new(
                new Zoom
                {
                    Percent = "100"
                },
                new Compatibility(
                    new CompatibilitySetting
                    {
                        Name = CompatSettingNameValues.CompatibilityMode,
                        Uri = "http://schemas.microsoft.com/office/word",
                        Val = "15"
                    }));

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
