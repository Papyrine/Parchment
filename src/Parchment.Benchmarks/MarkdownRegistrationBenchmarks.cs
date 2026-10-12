// Covers markdown registration, which nothing measured: RenderBenchmarks does register a markdown
// template, but inside [GlobalSetup], which BenchmarkDotNet runs untimed. Registration opens the
// style-source package twice — NormalizeStyleSource clones and parses it to settle the document
// type, then ScanNonBodyParts clones and parses it again to bind header/footer tokens — on top of
// the byte[] ToBytes has already produced from the caller's stream. Each benchmark below pins one
// branch of that path, so folding the two passes into one open has a before/after to point at.
// Watch Allocated as much as Mean: most of what is removable is full-package array copies, which at
// these template sizes sit inside the parse noise on the time columns.
[Config(typeof(BenchmarkConfig))]
public class MarkdownRegistrationBenchmarks
{
    byte[] docxStyleSource = null!;
    byte[] dotxStyleSource = null!;
    byte[] headerTokenStyleSource = null!;

    [GlobalSetup]
    public void Setup()
    {
        docxStyleSource = BuildStyleSource(WordprocessingDocumentType.Document, headerText: null);
        dotxStyleSource = BuildStyleSource(WordprocessingDocumentType.Template, headerText: null);
        headerTokenStyleSource = BuildStyleSource(WordprocessingDocumentType.Document, "Report: {{ Report.Title }}");
    }

    // No style source: the cached BlankDocxTemplate flows through both passes untouched, so each
    // returns the input array and never serializes its clone. The cheapest path — and the one a
    // merge has to avoid making more expensive.
    [Benchmark]
    public void NoStyleSource()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(markdownSource);
    }

    // The common case: a real .docx style source with nothing to rewrite. Both passes clone and
    // fully parse the package, then discard the clone and hand back the bytes they were given.
    [Benchmark]
    public void DocxStyleSource()
    {
        using var stream = new MemoryStream(docxStyleSource);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(markdownSource, stream);
    }

    // The same package behind a non-MemoryStream, so ToBytes takes its CopyTo-then-ToArray path
    // (two copies) rather than the single ToArray.
    [Benchmark]
    public void BufferedStyleSource()
    {
        using var memory = new MemoryStream(docxStyleSource);
        using var buffered = new BufferedStream(memory);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(markdownSource, buffered);
    }

    // A .dotx style source — the worst case. NormalizeStyleSource retypes, saves and serializes,
    // and then ScanNonBodyParts clones and parses those retyped bytes all over again.
    [Benchmark]
    public void DotxStyleSource()
    {
        using var stream = new MemoryStream(dotxStyleSource);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(markdownSource, stream);
    }

    // A header carrying a token, so the scan finds something to bind: anchors are baked in, and the
    // pass saves, extracts the scope trees and serializes instead of returning the input bytes.
    [Benchmark]
    public void HeaderTokenStyleSource()
    {
        using var stream = new MemoryStream(headerTokenStyleSource);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(markdownSource, stream);
    }

    readonly string markdownSource =
        """
        # {{ Report.Title }}

        *{{ Report.Author }} — {{ Report.Date }}*

        {{ Report.Summary }}
        """;

    static byte[] BuildStyleSource(WordprocessingDocumentType type, string? headerText)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, type))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(new Paragraph()));

            if (headerText != null)
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>("rIdH1");
                headerPart.Header = new(Para(headerText));
                mainPart.Document.Body!.Append(
                    new SectionProperties(
                        new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdH1"
                        }));
            }

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            styles.Append(NamedStyle("Normal", "Normal", isDefault: true));
            for (var i = 1; i <= 6; i++)
            {
                styles.Append(NamedStyle($"Heading{i}", $"Heading{i}"));
            }

            styles.Append(NamedStyle("ListParagraph", "List Paragraph"));
            styles.Append(NamedStyle("Quote", "Quote"));
            stylesPart.Styles = styles;
        }

        return stream.ToArray();
    }

    static Style NamedStyle(string styleId, string name, bool isDefault = false)
    {
        var style = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId,
            StyleName = new()
            {
                Val = name
            }
        };

        if (isDefault)
        {
            style.Default = true;
        }

        return style;
    }

    static Paragraph Para(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
