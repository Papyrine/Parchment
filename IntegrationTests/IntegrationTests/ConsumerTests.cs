public partial class ConsumerTests
{
    public class Customer
    {
        public required string Name { get; init; }
    }

    [ParchmentBindable]
    public partial class InvoiceModel
    {
        public required string Number { get; init; }
        public required Customer Customer { get; init; }
    }

    [Test]
    public async Task RegisterAndRender()
    {
        using var template = BuildTemplate();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<InvoiceModel>(template);

        using var stream = new MemoryStream();
        await store.Render(new InvoiceModel
        {
            Number = "INT-001",
            Customer = new() { Name = "Acme" }
        }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task LoopsAndConditionals()
    {
        using var template = BuildLoopTemplate();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<LoopModel>(template);

        using var stream = new MemoryStream();
        await store.Render(new LoopModel
        {
            Items = ["alpha", "beta", "gamma"]
        }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [ParchmentBindable]
    public partial class LoopModel
    {
        public required IReadOnlyList<string> Items { get; init; }
    }

    static MemoryStream BuildLoopTemplate()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("{% for item in Items %}") { Space = SpaceProcessingModeValues.Preserve })),
                new Paragraph(new Run(new Text("- {{ item }}") { Space = SpaceProcessingModeValues.Preserve })),
                new Paragraph(new Run(new Text("{% endfor %}") { Space = SpaceProcessingModeValues.Preserve }))));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }.AppendChild(new StyleName { Val = "Normal" }).Parent!);
            stylesPart.Styles = styles;
        }

        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildTemplate()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Invoice {{ Number }}") { Space = SpaceProcessingModeValues.Preserve })),
                new Paragraph(new Run(new Text("Customer: {{ Customer.Name }}") { Space = SpaceProcessingModeValues.Preserve }))));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            styles.Append(new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }.AppendChild(new StyleName { Val = "Normal" }).Parent!);
            stylesPart.Styles = styles;
        }

        stream.Position = 0;
        return stream;
    }

    // Run on demand to (re)generate the SG template fixtures committed alongside this file.
    // Excluded from default `dotnet run` runs by [Explicit]; invoke with:
    //   dotnet run --project IntegrationTests/IntegrationTests --configuration Release \
    //     -- --treenode-filter "/*/*/ConsumerTests/GenerateSgTemplates"
    [Test, Explicit]
    public async Task GenerateSgTemplates()
    {
        var dir = Path.GetDirectoryName(SourcePath())!;

        var docxPath = Path.Combine(dir, "sg-template.docx");
        await using (var fs = File.Create(docxPath))
        {
            using var doc = WordprocessingDocument.Create(fs, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Hello {{ Customer.Name }}") { Space = SpaceProcessingModeValues.Preserve })),
                new Paragraph(new Run(new Text("Invoice {{ Number }}") { Space = SpaceProcessingModeValues.Preserve }))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }.AppendChild(new StyleName { Val = "Normal" }).Parent!);
        }

        var mdPath = Path.Combine(dir, "sg-template.md");
        await File.WriteAllTextAsync(
            mdPath,
            """
            # Report for {{ Customer.Name }}

            {% for line in Lines %}
            - {{ line.Description }}
            {% endfor %}
            """);
    }

    static string SourcePath([System.Runtime.CompilerServices.CallerFilePath] string path = "") => path;
}
