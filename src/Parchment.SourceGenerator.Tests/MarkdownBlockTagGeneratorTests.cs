/// <summary>
/// Generator counterpart of the runtime <c>MarkdownBlockTagTests</c>: the same
/// <see cref="MarkdownBlockTagCases"/>, with the same accept or reject outcome.
/// </summary>
public class MarkdownBlockTagGeneratorTests
{
    const string model =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Name { get; set; } = "";
            public List<Line> Children { get; set; } = new();
        }

        [ParchmentModel]
        public partial class Doc
        {
            public string Name { get; set; } = "";
            public bool Flag { get; set; }
            public bool Off { get; set; }
            public int Count { get; set; }
            public List<Line> Lines { get; set; } = new();
        }
        """;

    [Test]
    [MethodDataSource(typeof(MarkdownBlockTagCases), nameof(MarkdownBlockTagCases.Accepted))]
    public async Task NoDiagnostic(string template) =>
        await Assert.That(Diagnostics(template).Select(_ => _.Id)).IsEmpty();

    [Test]
    [MethodDataSource(typeof(MarkdownBlockTagCases), nameof(MarkdownBlockTagCases.Rejected))]
    public async Task ReportsMissingMember(string template, string path)
    {
        var diagnostics = Diagnostics(template);
        await Assert.That(diagnostics.Select(_ => _.Id)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(diagnostics[0].GetMessage()).Contains($"'{path}'");
    }

    [Test]
    public async Task TableRowSourceThatIsNotEnumerable()
    {
        var diagnostics = Diagnostics("<table>{% tablerow c in Count %}x{% endtablerow %}</table>");
        await Assert.That(diagnostics.Select(_ => _.Id)).IsEquivalentTo(["PARCH002"]);
        await Assert.That(diagnostics[0].GetMessage()).Contains("{% tablerow c in ... %}");
    }

    static List<Diagnostic> Diagnostics(string markdown) =>
        GeneratorDriver.RunMarkdown(model, markdown)
            .Results.Single()
            .Diagnostics
            .ToList();
}
