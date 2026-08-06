/// <summary>
/// How a <c>[ParchmentModel]</c> type finds its files by convention: the template is the
/// AdditionalFile named <c>TypeName.parchment.md</c> or <c>TypeName.parchment.docx</c>, and a
/// markdown template's style source is the <c>TypeName.parchment.dotx</c> when the project carries
/// one, otherwise the nearest <c>parchment.dotx</c> up the directory tree.
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
        var result = Run(new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    // The convention matches by file name wherever the file sits — templates in a Templates/
    // folder need no path in the attribute.
    [Test]
    public async Task TemplateInANestedFolder()
    {
        var result = Run(new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task MatchIsCaseInsensitive()
    {
        var result = Run(new TemplateFile("report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task NoMatchIsReported()
    {
        var result = Run(new TemplateFile("Other.parchment.md", Markdown("Hello")));

        await Assert.That(Codes(result)).Contains("PARCH004");
    }

    // The marker is the whole reason the package can glob templates in without a csproj entry:
    // an unmarked file is not a template, however it is named and however it is declared.
    [Test]
    public async Task UnmarkedFileNamedAfterTheTypeDoesNotBind()
    {
        var result = Run(new TemplateFile("Report.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH004");
    }

    // The case that motivated the marker: a design note sitting beside the template, named after
    // the model because it is about the model. It must not become a second candidate.
    [Test]
    public async Task UnmarkedNamesakeIsNotAmbiguousWithTheTemplate()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Docs/Report.md", Markdown("Notes about the report model.")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    // An unmarked dotx is not a style document either — including one named after the type.
    [Test]
    public async Task UnmarkedDotxIsNotAStyleSource()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var result = Run(
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Report.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // The shared style document keeps its bare name — parchment.dotx, not parchment.parchment.dotx
    // — so a model actually named Parchment is the one case where the two could be confused. The
    // marker keeps them apart.
    [Test]
    public async Task ModelNamedParchmentDoesNotBindTheSharedStyleDoc()
    {
        var sharedBytes = GeneratorDriver.BuildDocxBytes("shared");
        // Namespaced: a type named Parchment in the global namespace would make `using Parchment;`
        // a CS0138, and the model would never reach the generator to prove anything.
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class Parchment
            {
                public required string Name { get; init; }
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(
            source,
            new TemplateFile("Templates/Parchment.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/parchment.dotx", sharedBytes));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();

        // Not PARCH021: the shared document is not a TypeName match, it is the folder's default.
        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(sharedBytes));
    }

    // Two namesakes in different folders: neither is preferred.
    [Test]
    public async Task NamesakesInDifferentFoldersAreAmbiguous()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Drafts/Report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH020");
    }

    [Test]
    public async Task TypeNamedDotxIsTheStyleSource()
    {
        var result = Run(
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).IsEmpty();
        // The style source is embedded as the second registration argument.
        await Assert.That(Generated(result)).Contains("RegisterMarkdownTemplate");
        await Assert.That(Generated(result)).Contains("FromBase64String");
    }

    [Test]
    public async Task TwoTypeNamedDotxesAreAmbiguous()
    {
        var result = Run(
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")),
            new TemplateFile("Drafts/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).Contains("PARCH021");
    }

    // parchment.dotx beside the template applies when no TypeName.dotx exists.
    [Test]
    public async Task SharedParchmentDotxApplies()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
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
            new TemplateFile("Templates/Inner/Report.parchment.md", Markdown("Hello {{ Name }}")),
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
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
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
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.parchment.dotx", typeBytes),
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
            new TemplateFile("Report.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }}")),
            new TemplateFile("Report.parchment.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains("RegisterDocxTemplate");
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // The embedded template is the file's own bytes, byte for byte.
    [Test]
    public async Task DocxTemplateBytesAreEmbedded()
    {
        var templateBytes = GeneratorDriver.BuildDocxBytes("Hello {{ Name }}");
        var result = Run(new TemplateFile("Report.parchment.docx", templateBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(templateBytes));
    }
}
