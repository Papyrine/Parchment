/// <summary>
/// PARCH034, and the string-list classification that shares its definition of a sequence.
/// <c>[ExcelsiorTable]</c> on a member Excelsior cannot tabulate used to be dropped without a
/// diagnostic, in every flow; a <c>string[,]</c> used to be classified as a string list the runtime
/// then read no items from.
/// </summary>
public class ExcelsiorTableShapeTests
{
    const string preamble =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        public enum Kind { A, B }

        """;

    static string Model(string attribute, string member) =>
        preamble +
        $$"""
          [{{attribute}}]
          public partial class Invoice
          {
              public string Title { get; set; } = "";
              {{member}}
          }
          """;

    static GeneratorDriverRunResult RunBindable(string source)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }

    static string Generated(GeneratorDriverRunResult result) =>
        string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

    [Test]
    [Arguments("public Line Owner { get; set; } = new();", "Sample.Line")]
    [Arguments("public string Text { get; set; } = \"\";", "string")]
    [Arguments("public List<string> Tags { get; set; } = new();", "List<string>")]
    [Arguments("public List<int> Numbers { get; set; } = new();", "List<int>")]
    [Arguments("public List<int?> Maybe { get; set; } = new();", "List<int?>")]
    [Arguments("public List<Kind> Kinds { get; set; } = new();", "Sample.Kind>")]
    [Arguments("public List<object> Things { get; set; } = new();", "List<object>")]
    [Arguments("public List<dynamic> Dyn { get; set; } = new();", "List<dynamic>")]
    [Arguments("public System.Collections.IEnumerable Untyped { get; set; } = new List<Line>();", "IEnumerable")]
    [Arguments("public Line[,] Grid { get; set; } = new Line[0, 0];", "Sample.Line[,]")]
    public async Task NotTabular_ReportsInEveryFlow(string member, string typeDisplay)
    {
        var name = member.Split(' ')[2];
        var declaration = $"[ExcelsiorTable] {member}";

        var docx = GeneratorDriver.Run(Model("ParchmentModel", declaration), "{{ Title }}");
        var markdown = GeneratorDriver.RunMarkdown(Model("ParchmentModel", declaration), "# {{ Title }}");
        var bindable = RunBindable(Model("ParchmentBindable", declaration));

        foreach (var result in new[] { docx, markdown, bindable })
        {
            var error = result.Results.Single().Diagnostics.SingleOrDefault(_ => _.Id == "PARCH034");
            await Assert.That(error).IsNotNull();
            await Assert.That(error!.Severity).IsEqualTo(DiagnosticSeverity.Error);
            await Assert.That(error.GetMessage()).Contains($"'Sample.Invoice.{name}'");
            await Assert.That(error.GetMessage()).Contains(typeDisplay);

            // Not emitted as a table, and not quietly reclassified as a string list either: the
            // diagnostic is the only thing the member produces.
            var generated = Generated(result);
            await Assert.That(generated).DoesNotContain("_ExcelsiorTables");
            await Assert.That(generated).DoesNotContain("_StringLists");
        }
    }

    [Test]
    [Arguments("public List<Line> Lines { get; set; } = new();")]
    [Arguments("public Line[] Lines { get; set; } = [];")]
    [Arguments("public IEnumerable<Line> Lines { get; set; } = [];")]
    [Arguments("public IReadOnlyList<Line?> Lines { get; set; } = [];")]
    [Arguments("public Line[][] Lines { get; set; } = [];")]
    // Excelsior reads fields as well as properties, so a tuple's Item1 / Item2 are columns.
    [Arguments("public List<(int A, string B)> Lines { get; set; } = new();")]
    [Arguments("public Dictionary<string, Line> Lines { get; set; } = new();")]
    public async Task Tabular_StaysSilentAndIsEmitted(string member)
    {
        var declaration = $"[ExcelsiorTable] {member}";

        var docx = GeneratorDriver.Run(Model("ParchmentModel", declaration), "{{ Lines }}");
        var markdown = GeneratorDriver.RunMarkdown(Model("ParchmentModel", declaration), "{{ Lines }}");
        var bindable = RunBindable(Model("ParchmentBindable", declaration));

        foreach (var result in new[] { docx, markdown, bindable })
        {
            await Assert.That(result.Results.Single().Diagnostics.Where(_ => _.Id == "PARCH034")).IsEmpty();
            await Assert.That(Generated(result)).Contains("new(\"Lines\", typeof(");
        }
    }

    // A static member already has PARCH019 saying the attribute does nothing; a second diagnostic
    // about the type of a member that is not dispatched at all would be noise.
    [Test]
    public async Task StaticMember_IsLeftToParch019()
    {
        var source = Model("ParchmentModel", "[ExcelsiorTable] public static string Text { get; set; } = \"\";");
        var diagnostics = GeneratorDriver.Run(source, "{{ Title }}").Results.Single().Diagnostics;

        await Assert.That(diagnostics.Where(_ => _.Id == "PARCH034")).IsEmpty();
        await Assert.That(diagnostics.Where(_ => _.Id == "PARCH019")).HasSingleItem();
    }

    [Test]
    public async Task NestedMember_IsReportedAgainstItsDeclaringType()
    {
        var source =
            preamble +
            """
            public class Customer
            {
                [ExcelsiorTable]
                public Line Primary { get; set; } = new();
            }

            [ParchmentModel]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """;
        var diagnostics = GeneratorDriver.Run(source, "{{ Customer.Primary.Description }}").Results.Single().Diagnostics;

        var error = diagnostics.Single(_ => _.Id == "PARCH034");
        await Assert.That(error.GetMessage()).Contains("'Sample.Customer.Primary'");
    }

    [Test]
    [Arguments("string[]")]
    [Arguments("List<string>")]
    [Arguments("IEnumerable<string>")]
    [Arguments("IReadOnlyList<string>")]
    [Arguments("IList<string?>")]
    [Arguments("HashSet<string>")]
    [Arguments("System.Collections.Immutable.ImmutableArray<string>")]
    public async Task StringSequence_IsAStringList(string type)
    {
        var source = Model("ParchmentBindable", $"public {type} Tags {{ get; set; }} = default!;");
        var generated = Generated(RunBindable(source));

        await Assert.That(generated).Contains("_StringLists");
        await Assert.That(generated).Contains("new(\"Tags\", o => ((global::Sample.Invoice)o).Tags)");
    }

    // The runtime reads a string list through IEnumerable<string>. A rank-2 array is not one, so
    // classifying it here replaced the token's paragraph with no items at all. A jagged array is a
    // sequence of string[], never of string; a string is a sequence of char.
    [Test]
    [Arguments("string[,]")]
    [Arguments("string[,,]")]
    [Arguments("string[][]")]
    [Arguments("List<string[]>")]
    [Arguments("string")]
    [Arguments("int[]")]
    [Arguments("System.Collections.IEnumerable")]
    public async Task AnythingElse_IsNotAStringList(string type)
    {
        var source = Model("ParchmentBindable", $"public {type} Tags {{ get; set; }} = default!;");
        var result = RunBindable(source);

        await Assert.That(Generated(result)).DoesNotContain("_StringLists");
        await Assert.That(result.Results.Single().Diagnostics).IsEmpty();
    }
}
