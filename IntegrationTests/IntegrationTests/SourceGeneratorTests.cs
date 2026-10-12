// End-to-end test of the [ParchmentModel] source generator at the packed-nupkg boundary.
// The unit tests in src/Parchment.SourceGenerator.Tests run the generator via a hand-rolled
// CSharpGeneratorDriver — they prove the generator's logic is correct, but they don't catch
// packaging regressions (analyzer path inside the nupkg, PackageShader sibling DLL packing,
// IsRoslynComponent flag, etc.). This project consumes Parchment as a real PackageReference,
// so a broken nupkg surfaces here as a compile-time failure.
//
// The models are intentionally **nested** inside the test class to exercise the nested-class
// emission path (PARCH011 / partial-enclosing wrapping). A previous version of the SG would
// silently emit a top-level partial that didn't combine with the nested target — caught here.

namespace IntegrationTests.Sg;

public class SgCustomer
{
    public required string Name { get; init; }
}

public class SgLine
{
    public required string Description { get; init; }
}

public partial class SourceGeneratorTests
{
    [ParchmentModel]
    public partial class SgInvoiceModel
    {
        public required string Number { get; init; }
        public required SgCustomer Customer { get; init; }
    }

    [ParchmentModel]
    public partial class SgReportModel
    {
        public required SgCustomer Customer { get; init; }
        public required IReadOnlyList<SgLine> Lines { get; init; }
    }

    // No registration call — the generated module initializer stored the embedded template when
    // this assembly loaded, and the store materializes it on first render.
    [Test]
    public async Task DocxTemplate_RendersFromTheGeneratedDefinition()
    {
        var store = new TemplateStore();

        using var stream = new MemoryStream();
        await store.Render(
            new SgInvoiceModel
            {
                Number = "SG-001",
                Customer = new()
                {
                    Name = "Acme"
                }
            },
            stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task MarkdownTemplate_RendersFromTheGeneratedDefinition()
    {
        var store = new TemplateStore();

        using var stream = new MemoryStream();
        await store.Render(
            new SgReportModel
            {
                Customer = new()
                {
                    Name = "Acme"
                },
                Lines =
                [
                    new()
                    {
                        Description = "Widget"
                    },
                    new()
                    {
                        Description = "Sprocket"
                    }
                ]
            },
            stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }
}
