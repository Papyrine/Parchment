/// <summary>
/// How the path in <c>[ParchmentModel]</c> finds its template: relative to the file the attribute
/// sits in, or as the tail of a declared template's path.
/// </summary>
public class TemplatePathResolutionTests
{
    // The driver writes the model source and the template into the same temp directory, so a path
    // relative to the model's file is a bare file name.
    static GeneratorDriverRunResult Run(string templatePath) =>
        GeneratorDriver.RunWithModelFile(
            $$"""
              using Parchment;

              [ParchmentModel("{{templatePath}}")]
              public partial class Report
              {
                  public required string Name { get; init; }
              }
              """,
            new TemplateFile("report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")));

    static IEnumerable<string> Codes(GeneratorDriverRunResult result) =>
        result.Diagnostics.Select(_ => _.Id);

    [Test]
    public async Task BareFileNameBesideTheModel()
    {
        var result = Run("report.md");

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task ExplicitlyRelativeToTheModelsDirectory()
    {
        var result = Run("./report.md");

        await Assert.That(Codes(result)).IsEmpty();
    }

    // Up and back down again lands on the same file, which is the form that proves the "." and ".."
    // segments are collapsed rather than compared as text.
    [Test]
    public async Task RelativePathThatClimbsAndReturns()
    {
        var result = Run("./sub/../report.md");

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task PathThatMatchesNothingIsReported()
    {
        var result = Run("./Templates/report.md");

        await Assert.That(Codes(result)).Contains("PARCH004");
    }

    // TemplatePath is what RegisterWith combines with a base path to find the template beside the
    // assembly, so it has to be where the file ends up rather than what the attribute said. A model
    // in Stocktakes/ naming "Templates/report.md" would otherwise send RegisterWith to a folder
    // that exists only in the source tree.
    [Test]
    public async Task TemplatePathIsWhereTheTemplateLandsNotWhatTheAttributeSaid()
    {
        var result = GeneratorDriver.RunInProject(
            """
            using Parchment;

            [ParchmentModel("Templates/report.md")]
            public partial class Report
            {
                public required string Name { get; init; }
            }
            """,
            "Stocktakes/Model.cs",
            new TemplateFile("Stocktakes/Templates/report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")));

        await Assert.That(Generated(result)).Contains("""TemplatePath => "Stocktakes/Templates/report.md";""");
    }

    [Test]
    public async Task TemplatePathIsUnchangedWhenTheAttributeAlreadyNamesTheRuntimeLocation()
    {
        var result = GeneratorDriver.RunInProject(
            """
            using Parchment;

            [ParchmentModel("Stocktakes/Templates/report.md")]
            public partial class Report
            {
                public required string Name { get; init; }
            }
            """,
            "Stocktakes/Model.cs",
            new TemplateFile("Stocktakes/Templates/report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")));

        await Assert.That(Generated(result)).Contains("""TemplatePath => "Stocktakes/Templates/report.md";""");
    }

    static string Generated(GeneratorDriverRunResult result) =>
        string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

    // Two templates whose paths both satisfy the attribute: one sitting beside the model, one whose
    // path merely ends the same way. Neither reading is wrong, so neither is chosen.
    [Test]
    public async Task AmbiguousPathIsReportedRatherThanPicked()
    {
        var result = GeneratorDriver.RunWithModelFile(
            """
            using Parchment;

            [ParchmentModel("Templates/report.md")]
            public partial class Report
            {
                public required string Name { get; init; }
            }
            """,
            new TemplateFile("Templates/report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")),
            new TemplateFile("Other/Templates/report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH020");
    }
}
