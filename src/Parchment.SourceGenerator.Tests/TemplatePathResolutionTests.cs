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
            ("report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")));

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
            ("Templates/report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")),
            ("Other/Templates/report.md", Encoding.UTF8.GetBytes("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH020");
    }
}
