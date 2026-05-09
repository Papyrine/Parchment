// End-to-end test of the [ParchmentTemplate] source generator at the packed-nupkg boundary.
// The unit tests in src/Parchment.SourceGenerator.Tests run the generator via a hand-rolled
// CSharpGeneratorDriver — they prove the generator's logic is correct, but they don't catch
// packaging regressions (analyzer path inside the nupkg, PackageShader sibling DLL packing,
// IsRoslynComponent flag, etc.). This project consumes Parchment as a real PackageReference,
// so a broken nupkg surfaces here as a compile-time failure.
//
// The fixtures are intentionally **nested** inside the test class to exercise the nested-class
// emission path (PARCH011 / partial-enclosing wrapping). A previous version of the SG would
// silently emit a top-level partial that didn't combine with the nested target — caught here.

namespace IntegrationTests.Sg;

public class SgCustomer
{
    public required string Name { get; init; }
}

public class SgInvoiceModel
{
    public required string Number { get; init; }
    public required SgCustomer Customer { get; init; }
}

public class SgLine
{
    public required string Description { get; init; }
}

public class SgReportModel
{
    public required SgCustomer Customer { get; init; }
    public required IReadOnlyList<SgLine> Lines { get; init; }
}

public partial class SourceGeneratorTests
{
    [ParchmentTemplate("sg-template.docx", typeof(SgInvoiceModel))]
    public partial class DocxFixture;

    [ParchmentTemplate("sg-template.md", typeof(SgReportModel))]
    public partial class MarkdownFixture;

    [Test]
    public async Task DocxTemplate_RegisterWithAndRender()
    {
        var store = new TemplateStore();
        DocxFixture.RegisterWith(store, basePath: AppContext.BaseDirectory);

        using var stream = new MemoryStream();
        await store.Render(
            DocxFixture.TemplateName,
            new SgInvoiceModel
            {
                Number = "SG-001",
                Customer = new() { Name = "Acme" }
            },
            stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task MarkdownTemplate_RegisterWithAndRender()
    {
        var store = new TemplateStore();
        MarkdownFixture.RegisterWith(store, basePath: AppContext.BaseDirectory);

        using var stream = new MemoryStream();
        await store.Render(
            MarkdownFixture.TemplateName,
            new SgReportModel
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
