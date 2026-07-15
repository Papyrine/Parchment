// Exercises Fix 4 (per-render part lookup). A template with several header/footer parts used to
// re-enumerate every part — re-allocating each part's Uri.ToString() — once per part looked up and
// again for the strip pass (O(parts²) plus redundant string allocs). The parts are now materialized
// into a dictionary once. Also covers Fix 3 (compatibilityMode is baked into the registration
// snapshot, so no per-render settings-part scan) and Fix 5 on every render. The absolute win is
// small at typical part counts — this mainly guards the multi-part path against regressions.
[Config(typeof(BenchmarkConfig))]
public class MultiPartRenderBenchmarks
{
    TemplateStore store = null!;
    HeaderFooterModel model = null!;

    [GlobalSetup]
    public void Setup()
    {
        model = new()
        {
            Title = "Annual Report",
            Author = "Ada Lovelace"
        };

        store = new();
        using var template = BuildTemplate();
        store.RegisterDocxTemplate<HeaderFooterModel>("multipart", template);
    }

    [Benchmark]
    public async Task RenderMultiPart()
    {
        using var output = new MemoryStream();
        await store.Render("multipart", model, output);
    }

    static MemoryStream BuildTemplate()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            AddHeader(mainPart, "rIdH1", "Header default: {{ Title }}");
            AddHeader(mainPart, "rIdH2", "Header first: {{ Title }}");
            AddHeader(mainPart, "rIdH3", "Header even: {{ Author }}");
            AddFooter(mainPart, "rIdF1", "Footer default: {{ Author }}");
            AddFooter(mainPart, "rIdF2", "Footer first: {{ Title }}");
            AddFooter(mainPart, "rIdF3", "Footer even: {{ Author }}");

            body.Append(Para("Body: {{ Title }} by {{ Author }}"));
            body.Append(
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdH1"
                    },
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.First,
                        Id = "rIdH2"
                    },
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Even,
                        Id = "rIdH3"
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdF1"
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.First,
                        Id = "rIdF2"
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Even,
                        Id = "rIdF3"
                    },
                    new PageSize
                    {
                        Width = 12240,
                        Height = 15840
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

        stream.Position = 0;
        return stream;
    }

    static void AddHeader(MainDocumentPart mainPart, string relationshipId, string text)
    {
        var headerPart = mainPart.AddNewPart<HeaderPart>(relationshipId);
        headerPart.Header = new(Para(text));
    }

    static void AddFooter(MainDocumentPart mainPart, string relationshipId, string text)
    {
        var footerPart = mainPart.AddNewPart<FooterPart>(relationshipId);
        footerPart.Footer = new(Para(text));
    }

    static Paragraph Para(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    public class HeaderFooterModel
    {
        public required string Title { get; init; }
        public required string Author { get; init; }
    }
}
