// PARCH028 — the compile-time half of the runtime's "block tags must share a container" check
// (BlockTagContainerTests in Parchment.Tests). A [ParchmentModel] template is only materialized
// on the model's first render, so without the diagnostic the mistake would get that far.
public class BlockTagContainerGeneratorTests
{
    const string source =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Invoice
        {
            public bool Flag { get; set; }

            public List<Line> Lines { get; set; } = new();
        }
        """;

    [Test]
    public async Task ForAndEndForInDifferentTableRows_ReportsParch028()
    {
        var diagnostics = Run(
            new Table(
                new TableRow(new TableCell(P("{% for line in Lines %}"))),
                new TableRow(new TableCell(P("{{ line.Description }}"))),
                new TableRow(new TableCell(P("{% endfor %}")))));

        var diagnostic = diagnostics.Single(_ => _.Id == "PARCH028");
        var message = diagnostic.GetMessage();
        await Assert.That(diagnostic.Severity).IsEqualTo(DiagnosticSeverity.Error);
        await Assert.That(message).Contains("{% for line in Lines %}");
        await Assert.That(message).Contains("{% endfor %}");
        await Assert.That(message).Contains("same container");
    }

    [Test]
    public async Task IfInCellAndEndIfInBody_ReportsParch028()
    {
        var diagnostics = Run(
            new Table(
                new TableRow(
                    new TableCell(
                        P("{% if Flag %}"),
                        P("In the cell")))),
            P("Between"),
            P("{% endif %}"));

        var message = diagnostics.Single(_ => _.Id == "PARCH028").GetMessage();
        await Assert.That(message).Contains("{% if Flag %}");
        await Assert.That(message).Contains("{% endif %}");
    }

    [Test]
    public async Task ElseInADifferentContainer_ReportsParch028()
    {
        var diagnostics = Run(
            P("{% if Flag %}"),
            P("Yes"),
            new Table(
                new TableRow(
                    new TableCell(
                        P("{% else %}")))),
            P("No"),
            P("{% endif %}"));

        // Once per block, naming the first tag that strays.
        var message = diagnostics.Single(_ => _.Id == "PARCH028").GetMessage();
        await Assert.That(message).Contains("{% if Flag %}");
        await Assert.That(message).Contains("{% else %}");
    }

    [Test]
    public async Task LoopInsideOneCell_NoDiagnostic()
    {
        var diagnostics = Run(
            new Table(
                new TableRow(
                    new TableCell(
                        P("{% for line in Lines %}"),
                        P("{{ line.Description }}"),
                        P("{% endfor %}")))));

        await Assert.That(diagnostics.Select(_ => _.Id)).IsEmpty();
    }

    [Test]
    public async Task LoopRepeatingAWholeTable_NoDiagnostic()
    {
        var diagnostics = Run(
            P("{% for line in Lines %}"),
            new Table(
                new TableRow(
                    new TableCell(
                        P("{{ line.Description }}")))),
            P("{% endfor %}"));

        await Assert.That(diagnostics.Select(_ => _.Id)).IsEmpty();
    }

    [Test]
    public async Task BlockInACellNestedInABodyLoop_NoDiagnostic()
    {
        var diagnostics = Run(
            P("{% for line in Lines %}"),
            new Table(
                new TableRow(
                    new TableCell(
                        P("{% if line.Description %}"),
                        P("{{ line.Description }}"),
                        P("{% endif %}")))),
            P("{% endfor %}"));

        await Assert.That(diagnostics.Select(_ => _.Id)).IsEmpty();
    }

    [Test]
    public async Task UnsupportedTag_LeavesNestingUnreported()
    {
        // With a tag the generator cannot classify, which tag closes which block is a guess, so
        // PARCH003 stands alone.
        var diagnostics = Run(
            P("{% for line in Lines %}"),
            new Table(
                new TableRow(
                    new TableCell(
                        P("{% unless Flag %}"),
                        P("{% endfor %}")))));

        var ids = diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH003");
        await Assert.That(ids).DoesNotContain("PARCH028");
    }

    static ImmutableArray<Diagnostic> Run(params OpenXmlElement[] blocks)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Invoice.parchment.docx", GeneratorDriver.BuildDocxBytesFromBlocks(blocks)));
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult().Results.Single().Diagnostics;
    }

    static Paragraph P(string text) =>
        GeneratorDriver.BuildParagraph(text);
}
