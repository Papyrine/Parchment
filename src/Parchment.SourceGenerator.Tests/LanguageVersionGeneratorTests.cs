/// <summary>
/// The generated file is compiled with the consumer's <c>LangVersion</c>, and C# 14 is the oldest
/// it is written for. A project below that gets PARCH038 on the model and nothing else: whatever
/// was emitted beside it would be compiled at the old version too.
/// </summary>
public class LanguageVersionGeneratorTests
{
    // Block namespaces and nothing newer than C# 8, so the model itself is legal at every version
    // tried and the only thing in question is the generator's output.
    const string model =
        """
        using Parchment;

        namespace Sample
        {
            [ParchmentModel]
            public partial class Order
            {
                public string Number { get; set; } = "";
            }
        }
        """;

    const string bindable =
        """
        using Parchment;

        namespace Sample
        {
            [ParchmentBindable]
            public partial class Bound
            {
                public string Number { get; set; } = "";
            }
        }
        """;

    static TemplateFile Docx(params string[] paragraphs) =>
        new("Order.parchment.docx", GeneratorDriver.BuildDocxBytes(paragraphs));

    [Test]
    [Arguments(LanguageVersion.CSharp8, "8.0")]
    [Arguments(LanguageVersion.CSharp9, "9.0")]
    [Arguments(LanguageVersion.CSharp10, "10.0")]
    [Arguments(LanguageVersion.CSharp13, "13.0")]
    public async Task BelowCSharp14_ReportsParch038_Docx(LanguageVersion version, string display)
    {
        var result = Run(model, version, Docx("{{ Number }}"));

        var diagnostic = result.Results.Single().Diagnostics.Single();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH038");
        await Assert.That(diagnostic.Severity).IsEqualTo(DiagnosticSeverity.Error);
        await Assert.That(diagnostic.GetMessage()).Contains("'Order'");
        await Assert.That(diagnostic.GetMessage()).Contains($"C# {display}.");
        // On the attribute that asked for generated code.
        var span = diagnostic.Location.SourceSpan;
        await Assert.That(model.Substring(span.Start, span.Length)).IsEqualTo("ParchmentModel");
        await Assert.That(result.GeneratedTrees).IsEmpty();
    }

    [Test]
    public async Task BelowCSharp14_ReportsParch038_Markdown()
    {
        var result = Run(
            model,
            LanguageVersion.CSharp13,
            new TemplateFile("Order.parchment.md", "# {{ Number }}\n"u8.ToArray()));

        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEquivalentTo(["PARCH038"]);
        await Assert.That(result.GeneratedTrees).IsEmpty();
    }

    [Test]
    public async Task BelowCSharp14_ReportsParch038_Bindable()
    {
        var result = Run(bindable, LanguageVersion.CSharp13);

        var diagnostic = result.Results.Single().Diagnostics.Single();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH038");
        await Assert.That(diagnostic.GetMessage()).Contains("'Bound'");
        await Assert.That(result.GeneratedTrees).IsEmpty();
    }

    // The version is the first thing to fix, and the only thing reported until it is: a missing
    // template or an unknown token would be noise beside it.
    [Test]
    public async Task BelowCSharp14_ReportsNothingElse()
    {
        var unknownToken = Run(model, LanguageVersion.CSharp13, Docx("{{ Missing }}"));
        var noTemplate = Run(model, LanguageVersion.CSharp13);

        await Assert.That(unknownToken.Results.Single().Diagnostics.Select(_ => _.Id).ToList())
            .IsEquivalentTo(["PARCH038"]);
        await Assert.That(noTemplate.Results.Single().Diagnostics.Select(_ => _.Id).ToList())
            .IsEquivalentTo(["PARCH038"]);
    }

    [Test]
    [Arguments(LanguageVersion.CSharp14)]
    [Arguments(LanguageVersion.Latest)]
    [Arguments(LanguageVersion.LatestMajor)]
    [Arguments(LanguageVersion.Default)]
    [Arguments(LanguageVersion.Preview)]
    public async Task CSharp14AndLater_Generates(LanguageVersion version)
    {
        var models = Run(model, version, Docx("{{ Number }}"));
        var bindables = Run(bindable, version);

        await Assert.That(models.Results.Single().Diagnostics).IsEmpty();
        await Assert.That(models.GeneratedTrees.Single().ToString()).Contains("namespace Sample;");
        await Assert.That(bindables.Results.Single().Diagnostics).IsEmpty();
        await Assert.That(bindables.GeneratedTrees.Length).IsEqualTo(1);
    }

    static GeneratorDriverRunResult Run(string source, LanguageVersion version, params TemplateFile[] files)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            files,
            modelFileName: null,
            languageVersion: version);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }
}
