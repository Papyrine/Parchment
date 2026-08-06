/// <summary>
/// How a <c>[ParchmentModel]</c> type finds its files by convention: the template is the
/// AdditionalFile named after the type, and a markdown template's style source is the
/// <c>TypeName.dotx</c> when the project carries one, otherwise the nearest
/// <c>parchment.dotx</c> up the directory tree.
/// </summary>
public class TemplateConventionTests
{
    const string reportModel =
        """
        using Parchment;

        [ParchmentModel]
        public partial class Report
        {
            public required string Name { get; init; }
        }
        """;

    static GeneratorDriverRunResult Run(params TemplateFile[] files)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(reportModel, files);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }

    static IEnumerable<string> Codes(GeneratorDriverRunResult result) =>
        result.Diagnostics.Select(_ => _.Id);

    static string Generated(GeneratorDriverRunResult result) =>
        string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

    static byte[] Markdown(string text) => Encoding.UTF8.GetBytes(text);

    [Test]
    public async Task TemplateBesideNothingElse()
    {
        var result = Run(new TemplateFile("Report.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    // The convention matches by file name wherever the file sits — templates in a Templates/
    // folder need no path in the attribute.
    [Test]
    public async Task TemplateInANestedFolder()
    {
        var result = Run(new TemplateFile("Templates/Report.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task MatchIsCaseInsensitive()
    {
        var result = Run(new TemplateFile("report.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task NoMatchIsReported()
    {
        var result = Run(new TemplateFile("Other.md", Markdown("Hello")));

        await Assert.That(Codes(result)).Contains("PARCH004");
    }

    // Two namesakes in different folders: neither is preferred.
    [Test]
    public async Task NamesakesInDifferentFoldersAreAmbiguous()
    {
        var result = Run(
            new TemplateFile("Templates/Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Drafts/Report.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH020");
    }

    [Test]
    public async Task TypeNamedDotxIsTheStyleSource()
    {
        var result = Run(
            new TemplateFile("Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Report.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).IsEmpty();
        // The style source is embedded as the second registration argument.
        await Assert.That(Generated(result)).Contains("RegisterMarkdownTemplate");
        await Assert.That(Generated(result)).Contains("FromBase64String");
    }

    [Test]
    public async Task TwoTypeNamedDotxesAreAmbiguous()
    {
        var result = Run(
            new TemplateFile("Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.dotx", GeneratorDriver.BuildDocxBytes("styled")),
            new TemplateFile("Drafts/Report.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).Contains("PARCH021");
    }

    // parchment.dotx beside the template applies when no TypeName.dotx exists.
    [Test]
    public async Task SharedParchmentDotxApplies()
    {
        var result = Run(
            new TemplateFile("Templates/Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains("FromBase64String");
    }

    // The nearest ancestor's parchment.dotx wins over one further up.
    [Test]
    public async Task NearestParchmentDotxWins()
    {
        var nearBytes = GeneratorDriver.BuildDocxBytes("near");
        var farBytes = GeneratorDriver.BuildDocxBytes("far");
        var result = Run(
            new TemplateFile("Templates/Inner/Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Inner/parchment.dotx", nearBytes),
            new TemplateFile("parchment.dotx", farBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(nearBytes));
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(farBytes));
    }

    // A parchment.dotx in a sibling folder is not on the template's ancestor path and must not
    // apply.
    [Test]
    public async Task SiblingFolderParchmentDotxDoesNotApply()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var result = Run(
            new TemplateFile("Templates/Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Other/parchment.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // TypeName.dotx wins over a shared parchment.dotx, even one beside the template.
    [Test]
    public async Task TypeNamedDotxWinsOverSharedParchmentDotx()
    {
        var typeBytes = GeneratorDriver.BuildDocxBytes("typed");
        var sharedBytes = GeneratorDriver.BuildDocxBytes("shared");
        var result = Run(
            new TemplateFile("Templates/Report.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.dotx", typeBytes),
            new TemplateFile("Templates/parchment.dotx", sharedBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(typeBytes));
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(sharedBytes));
    }

    // Docx templates carry their own styles; a dotx present in the project must not affect them.
    [Test]
    public async Task DocxTemplateIgnoresStyleDocs()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var result = Run(
            new TemplateFile("Report.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }}")),
            new TemplateFile("Report.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains("RegisterDocxTemplate");
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // The embedded template is the file's own bytes, byte for byte.
    [Test]
    public async Task DocxTemplateBytesAreEmbedded()
    {
        var templateBytes = GeneratorDriver.BuildDocxBytes("Hello {{ Name }}");
        var result = Run(new TemplateFile("Report.docx", templateBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(templateBytes));
    }
}
