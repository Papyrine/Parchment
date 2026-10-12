// The shape Word produces for an ordinary letterhead: three headers, three footers, footnotes and
// endnotes, none of them carrying a token, and a body that does. A render used to read every one of
// those parts, which made the SDK parse it and then write it back out on save. Only the parts with
// a scope tree are read now, so the rest stay as the bytes cloned from the registration snapshot.
[Config(typeof(BenchmarkConfig))]
public class TokenFreePartsRenderBenchmarks
{
    const int paragraphsPerPart = 20;

    TemplateStore store = null!;
    LetterModel model = null!;

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
        store.RegisterDocxTemplate<LetterModel>(template);
    }

    [Benchmark]
    public async Task RenderWithTokenFreeParts()
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
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            var section = new SectionProperties();
            foreach (var type in new[] { HeaderFooterValues.Default, HeaderFooterValues.First, HeaderFooterValues.Even })
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                headerPart.Header = new(Paragraphs("Header"));
                section.Append(
                    new HeaderReference
                    {
                        Type = type,
                        Id = mainPart.GetIdOfPart(headerPart)
                    });
            }

            foreach (var type in new[] { HeaderFooterValues.Default, HeaderFooterValues.First, HeaderFooterValues.Even })
            {
                var footerPart = mainPart.AddNewPart<FooterPart>();
                footerPart.Footer = new(Paragraphs("Footer"));
                section.Append(
                    new FooterReference
                    {
                        Type = type,
                        Id = mainPart.GetIdOfPart(footerPart)
                    });
            }

            mainPart.AddNewPart<FootnotesPart>().Footnotes = new(
                new Footnote(Paragraphs("Footnote"))
                {
                    Id = 1
                });
            mainPart.AddNewPart<EndnotesPart>().Endnotes = new(
                new Endnote(Paragraphs("Endnote"))
                {
                    Id = 1
                });

            body.Append(Para("{{ Title }} by {{ Author }}"));
            body.Append(Paragraphs("Body"));
            section.Append(
                new PageSize
                {
                    Width = 12240,
                    Height = 15840
                });
            body.Append(section);
        }

        stream.Position = 0;
        return stream;
    }

    static IEnumerable<OpenXmlElement> Paragraphs(string label)
    {
        for (var index = 0; index < paragraphsPerPart; index++)
        {
            yield return Para($"{label} line {index}: lorem ipsum dolor sit amet, consectetur adipiscing elit.");
        }
    }

    static Paragraph Para(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    public class LetterModel
    {
        public required string Title { get; init; }
        public required string Author { get; init; }
    }
}
