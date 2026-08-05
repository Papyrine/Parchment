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
/// The template travels inside the generated source — nothing is deployed beside the assembly and
/// nothing is embedded as a manifest resource.
/// </summary>
/// <remarks>
/// Kept as a test rather than only as documentation, because a regression here surfaces at a
/// consumer as a runtime failure on an app whose build was green.
/// </remarks>
public partial class EmbeddedTemplateTests
{
    [ParchmentModel]
    public partial class EmbeddedReportModel
    {
        public required EmbeddedCustomer Customer { get; init; }
        public required IReadOnlyList<EmbeddedLine> Lines { get; init; }
    }

    // The point of embedding into the generated source: nothing loose to deploy.
    [Test]
    public async Task TheTemplateIsNotCopiedBesideTheAssembly() =>
        await Assert.That(File.Exists(Path.Combine(AppContext.BaseDirectory, "EmbeddedReportModel.md"))).IsFalse();

    // And no manifest resource either — the old embedding mechanism is gone.
    [Test]
    public async Task TheTemplateIsNotAManifestResource() =>
        await Assert.That(
                typeof(EmbeddedTemplateTests).Assembly.GetManifestResourceNames()
                    .Any(_ => _.Contains("EmbeddedReportModel", StringComparison.OrdinalIgnoreCase)))
            .IsFalse();

    [Test]
    public async Task RendersWithNoFileOnDisk()
    {
        var store = new TemplateStore();

        using var stream = new MemoryStream();
        await store.Render(
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
