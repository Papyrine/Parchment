namespace IntegrationTests.Embedded;

public class EmbeddedCustomer
{
    public required string Name { get; init; }
}

public class EmbeddedLine
{
    public required string Description { get; init; }
}

/// <summary>
/// Consuming a template declared with <c>&lt;ParchmentEmbeddedTemplate&gt;</c>.
/// </summary>
/// <remarks>
/// The generated <c>RegisterWith</c> helper reads the template from disk, which an embedded
/// template deliberately is not — so registration goes through the manifest instead. Kept as a test
/// rather than only as documentation, because the two ways of getting the template to the store are
/// easy to confuse and the failure is a runtime <c>FileNotFoundException</c>.
/// </remarks>
public partial class EmbeddedTemplateTests
{
    [ParchmentModel("sg-embedded.md")]
    public partial class EmbeddedReportModel
    {
        public required EmbeddedCustomer Customer { get; init; }
        public required IReadOnlyList<EmbeddedLine> Lines { get; init; }
    }

    // The resource is named as though it had been embedded in place, so the staging under obj does
    // not leak into the name a caller has to know.
    const string ResourceName = "IntegrationTests.sg-embedded.md";

    static string ReadTemplate()
    {
        var assembly = typeof(EmbeddedTemplateTests).Assembly;
        using var stream = assembly.GetManifestResourceStream(ResourceName) ??
                           throw new InvalidOperationException(
                               $"'{ResourceName}' is not embedded. Found: {string.Join(", ", assembly.GetManifestResourceNames())}");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    [Test]
    public async Task TheTemplateIsEmbeddedUnderItsInPlaceName() =>
        await Assert.That(typeof(EmbeddedTemplateTests).Assembly.GetManifestResourceNames())
            .Contains(ResourceName);

    // The point of the item type: nothing loose to deploy.
    [Test]
    public async Task TheTemplateIsNotCopiedBesideTheAssembly() =>
        await Assert.That(File.Exists(Path.Combine(AppContext.BaseDirectory, "sg-embedded.md"))).IsFalse();

    [Test]
    public async Task RegisterFromTheManifestAndRender()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmbeddedReportModel>(
            EmbeddedReportModel.TemplateName,
            ReadTemplate());

        using var stream = new MemoryStream();
        await store.Render(
            EmbeddedReportModel.TemplateName,
            new EmbeddedReportModel
            {
                Customer = new() { Name = "Acme" },
                Lines =
                [
                    new() { Description = "Widget" },
                    new() { Description = "Sprocket" }
                ]
            },
            stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }
}
