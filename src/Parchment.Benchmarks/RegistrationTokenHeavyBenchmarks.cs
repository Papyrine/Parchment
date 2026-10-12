// Registration of a template where most paragraphs carry tokens: loops with modifiers, conditionals
// and several substitutions to a paragraph, in the body and in a header and a footer. Every token
// is regex-matched and parsed by Fluid as its part is scanned, so this is the cost that scales with
// how much of a template is bound. Each block is nine token paragraphs holding seventeen tokens;
// RegistrationScalingBenchmarks covers the one-token-per-paragraph shape.
[Config(typeof(BenchmarkConfig))]
public class RegistrationTokenHeavyBenchmarks
{
    byte[] templateBytes = null!;

    [Params(10, 50)]
    public int Blocks { get; set; }

    [GlobalSetup]
    public void Setup() =>
        templateBytes = BuildTemplate(Blocks);

    [Benchmark]
    public void Register()
    {
        using var stream = new MemoryStream(templateBytes);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(stream);
    }

    static byte[] BuildTemplate(int blocks)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            for (var i = 0; i < blocks; i++)
            {
                AppendBlock(body);
            }

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            var header = new Header();
            AppendBlock(header);
            headerPart.Header = header;

            var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter");
            var footer = new Footer();
            AppendBlock(footer);
            footerPart.Footer = footer;

            body.Append(
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    },
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdFooter"
                    },
                    new PageSize
                    {
                        Width = 12240,
                        Height = 15840
                    }));
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

    static void AppendBlock(OpenXmlCompositeElement parent) =>
        parent.Append(
            Para("Invoice {{ Number }} issued {{ IssueDate }} due {{ DueDate }}"),
            Para("Bill to {{ Customer.Name | upcase }} <{{ Customer.Email }}>, {{ Customer.Address }}"),
            Para("{% if Customer.IsPreferred and Customer.VatNumber != null %}"),
            Para("Preferred customer, VAT {{ Customer.VatNumber }}"),
            Para("{% endif %}"),
            Para("{% for line in Lines limit: 20 offset: 1 reversed %}"),
            Para("{{ line.Quantity }} x {{ line.Description }} at {{ line.UnitPrice }} = {{ line.LineTotal }}"),
            Para("{% endfor %}"),
            Para("Total {{ Total }} {{ Currency }}"),
            Para("Thank you for your business."));

    static Paragraph Para(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
