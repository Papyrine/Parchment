[Config(typeof(BenchmarkConfig))]
public class RegistrationBenchmarks
{
    byte[] templateBytes = null!;
    string templatePath = null!;

    [GlobalSetup]
    public void Setup()
    {
        templateBytes = BuildTemplate();
        templatePath = Path.GetTempFileName();
        File.WriteAllBytes(templatePath, templateBytes);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (File.Exists(templatePath))
        {
            File.Delete(templatePath);
        }
    }

    [Benchmark]
    public void RegisterFromMemoryStream()
    {
        using var stream = new MemoryStream(templateBytes);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(stream);
    }

    [Benchmark]
    public void RegisterFromBufferedStream()
    {
        using var ms = new MemoryStream(templateBytes);
        using var buffered = new BufferedStream(ms);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(buffered);
    }

    [Benchmark]
    public void RegisterFromFilePath()
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(templatePath);
    }

    [Benchmark]
    public void RegisterViaSourceGeneratorPath()
    {
        // Simulates what a SG-emitted module initializer does — pre-populates the per-type Fluid
        // accessors + the StringList map (for Tags), then calls RegisterDocxTemplate. The
        // runtime's reflection walks short-circuit on the cache hits.
        Parchment.Generated.GeneratedRegistration.RegisterFluidAccessors(
            typeof(Invoice), invoiceAccessors);
        Parchment.Generated.GeneratedRegistration.RegisterFluidAccessors(
            typeof(Customer), customerAccessors);
        Parchment.Generated.GeneratedRegistration.RegisterFluidAccessors(
            typeof(LineItem), lineItemAccessors);
        Parchment.Generated.GeneratedRegistration.RegisterStringList(
            typeof(Invoice), invoiceStringLists);

        using var stream = new MemoryStream(templateBytes);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(stream);
    }

    static KeyValuePair<string, Fluid.IMemberAccessor>[] invoiceAccessors =
    [
        new("Number", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Number)),
        new("IssueDate", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).IssueDate)),
        new("DueDate", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).DueDate)),
        new("Customer", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Customer)),
        new("Lines", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Lines)),
        new("Currency", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Currency)),
        new("Notes", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Notes)),
        new("Tags", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Tags)),
        new("Subtotal", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Subtotal)),
        new("Tax", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Tax)),
        new("Total", new Fluid.Accessors.DelegateAccessor((o, _) => ((Invoice)o).Total))
    ];

    static KeyValuePair<string, Fluid.IMemberAccessor>[] customerAccessors =
    [
        new("Name", new Fluid.Accessors.DelegateAccessor((o, _) => ((Customer)o).Name)),
        new("Email", new Fluid.Accessors.DelegateAccessor((o, _) => ((Customer)o).Email)),
        new("Address", new Fluid.Accessors.DelegateAccessor((o, _) => ((Customer)o).Address)),
        new("VatNumber", new Fluid.Accessors.DelegateAccessor((o, _) => ((Customer)o).VatNumber)),
        new("IsPreferred", new Fluid.Accessors.DelegateAccessor((o, _) => ((Customer)o).IsPreferred))
    ];

    static KeyValuePair<string, Fluid.IMemberAccessor>[] lineItemAccessors =
    [
        new("Description", new Fluid.Accessors.DelegateAccessor((o, _) => ((LineItem)o).Description)),
        new("Quantity", new Fluid.Accessors.DelegateAccessor((o, _) => ((LineItem)o).Quantity)),
        new("UnitPrice", new Fluid.Accessors.DelegateAccessor((o, _) => ((LineItem)o).UnitPrice)),
        new("LineTotal", new Fluid.Accessors.DelegateAccessor((o, _) => ((LineItem)o).LineTotal))
    ];

    static readonly Parchment.Generated.StringListMapEntry[] invoiceStringLists =
    [
        new("Tags", o => ((Invoice)o).Tags)
    ];

    static byte[] BuildTemplate()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(
                new Body(
                    new Paragraph(
                        new Run(
                            new Text("Invoice {{ Number }}")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            })),
                    new Paragraph(
                        new Run(
                            new Text("Customer: {{ Customer.Name }}")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            })),
                    new Paragraph(
                        new Run(
                            new Text("Total: {{ Total }} {{ Currency }}")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }))));

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
                        }).Parent!);
            stylesPart.Styles = styles;
        }

        return stream.ToArray();
    }
}
