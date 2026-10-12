// The models live in ParchmentSample, where they are decorated with [ParchmentModel] and their
// templates are registered as AdditionalFiles. That project is what compiles the readme's
// source-generator examples, so they cannot drift from the generator's actual output.
public partial class GeneratedRegistrationTests
{
    [Test]
    public async Task GeneratedModuleInitializerCoversBothFlows()
    {
        #region GeneratorRender

        // No registration call: the generated module initializer stored each model's embedded
        // template when the model assembly loaded, and the store picks it up on first render.
        var store = new TemplateStore();

        using var invoice = new MemoryStream();
        await store.Render(SampleData.Invoice(), invoice);

        #endregion

        invoice.Position = 0;

        using var report = new MemoryStream();
        await store.Render(SampleData.Report(), report);
        report.Position = 0;

        await Assert.That(Text(invoice)).Contains("Invoice INV-2026-0042");
        await Assert.That(Text(report)).Contains("Q2 Platform Health Review");
    }

    [ParchmentBindable]
    public partial class LazyDocxModel
    {
        public string Name { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class LazyMarkdownModel
    {
        public string Name { get; set; } = "";
    }

    static byte[] TemplateBytes(string paragraph)
    {
        using var template = DocxTemplateBuilder.Build(paragraph);
        using var copy = new MemoryStream();
        template.CopyTo(copy);
        return copy.ToArray();
    }

    // A definition holds its bytes as a factory, and registering one must not run it: the
    // template is decoded when a store first renders the model, and once per store.
    [Test]
    public async Task DocxFactoryRunsOnFirstRenderOnly()
    {
        var bytes = TemplateBytes("Hello {{ Name }}");
        var calls = 0;
        Parchment.Generated.GeneratedTemplateDefinitions.Add(
            typeof(LazyDocxModel),
            new Parchment.Generated.DocxTemplateDefinition(
                () =>
                {
                    calls++;
                    return bytes;
                },
                ProtectionMode.WhenEditable));

        await Assert.That(calls).IsEqualTo(0);

        var store = new TemplateStore();
        using var first = new MemoryStream();
        await store.Render(new LazyDocxModel { Name = "one" }, first);
        using var second = new MemoryStream();
        await store.Render(new LazyDocxModel { Name = "two" }, second);

        await Assert.That(calls).IsEqualTo(1);
        first.Position = 0;
        await Assert.That(Text(first)).Contains("Hello one");
    }

    [Test]
    public async Task MarkdownStyleFactoryRunsOnFirstRenderOnly()
    {
        var bytes = TemplateBytes("styled");
        var calls = 0;
        Parchment.Generated.GeneratedTemplateDefinitions.Add(
            typeof(LazyMarkdownModel),
            new Parchment.Generated.MarkdownTemplateDefinition(
                "Hello {{ Name }}",
                () =>
                {
                    calls++;
                    return bytes;
                }));

        await Assert.That(calls).IsEqualTo(0);

        var store = new TemplateStore();
        using var first = new MemoryStream();
        await store.Render(new LazyMarkdownModel { Name = "one" }, first);
        using var second = new MemoryStream();
        await store.Render(new LazyMarkdownModel { Name = "two" }, second);

        await Assert.That(calls).IsEqualTo(1);
        first.Position = 0;
        await Assert.That(Text(first)).Contains("Hello one");
    }

    [ParchmentBindable]
    public partial class UnreadModel
    {
        public string Name { get; set; } = "";
    }

    // The generated registration names the type whose [EmbeddedBytes] attribute holds the
    // template. The attribute is not looked at until a render needs it — shown here by a type
    // that has none, which registers without complaint.
    [Test]
    public async Task EmbeddedBytesAreNotReadAtRegistration()
    {
        Parchment.Generated.GeneratedRegistration.RegisterDocxTemplate(typeof(UnreadModel), typeof(UnreadModel));

        var store = new TemplateStore();
        using var stream = new MemoryStream();
        var exception = await Assert.That(() => store.Render(new UnreadModel(), stream))
            .Throws<InvalidOperationException>();
        await Assert.That(exception!.Message).Contains("[EmbeddedBytes]");
    }

    static string Text(Stream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
