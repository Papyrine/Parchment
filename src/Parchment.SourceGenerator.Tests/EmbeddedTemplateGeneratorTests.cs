/// <summary>
/// Registration emitted for a template declared with <c>&lt;ParchmentEmbeddedTemplate&gt;</c>, which
/// is never copied to disk and so has to be read out of the assembly manifest.
/// </summary>
public class EmbeddedTemplateGeneratorTests
{
    [Test]
    public Task Markdown_EmbeddedTemplate_RegistersFromTheManifest()
    {
        var source =
            """
            using Parchment;

            [ParchmentModel("template.md")]
            public partial class Report
            {
                public required string Name { get; init; }
            }
            """;

        var result = GeneratorDriver.RunEmbedded(
            source,
            "template.md",
            Encoding.UTF8.GetBytes("Hello {{ Name }}"),
            "MyProject.template.md");

        return Verify(result);
    }

    [Test]
    public Task Docx_EmbeddedTemplate_RegistersFromTheManifest()
    {
        var source =
            """
            using Parchment;

            [ParchmentModel("template.docx")]
            public partial class Invoice
            {
                public required string Number { get; init; }
            }
            """;

        var result = GeneratorDriver.RunEmbedded(
            source,
            "template.docx",
            GeneratorDriver.BuildDocxBytes("{{ Number }}"),
            "MyProject.template.docx");

        return Verify(result);
    }
}
