// Isolates Fix 1 (Anchors bookmark-id allocation). Registration inserts one anchor bookmark per
// token-bearing paragraph. The old id allocator re-walked the whole part's BookmarkStart
// descendants for every anchor — O(paragraphs²) — so time grew super-linearly with the token
// count. Seeding the max id once and incrementing makes it O(paragraphs). Watch the per-item cost
// stay flat across the params instead of climbing.
[Config(typeof(BenchmarkConfig))]
public class RegistrationScalingBenchmarks
{
    byte[] templateBytes = null!;

    [Params(10, 100, 500)]
    public int TokenParagraphs { get; set; }

    [GlobalSetup]
    public void Setup() =>
        templateBytes = BuildTemplate(TokenParagraphs);

    [Benchmark]
    public void Register()
    {
        using var stream = new MemoryStream(templateBytes);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>("bench", stream);
    }

    static byte[] BuildTemplate(int tokenParagraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            for (var i = 0; i < tokenParagraphs; i++)
            {
                body.AppendChild(Para("Invoice {{ Number }}"));
            }

            mainPart.Document = new(body);

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

    static Paragraph Para(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
