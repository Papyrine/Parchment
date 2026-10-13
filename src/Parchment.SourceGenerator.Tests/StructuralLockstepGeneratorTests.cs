// The compile-time half of StructuralLockstepCases (linked from Parchment.Tests, where
// StructuralLockstepTests registers the same templates and pins what the runtime says). A
// [ParchmentModel] template is only materialized on the model's first render, so a structural
// mistake the generator lets through gets that far.
public class StructuralLockstepGeneratorTests
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
        public partial class Model
        {
            public bool Flag { get; set; }

            public string Name { get; set; } = "";

            public List<Line> Lines { get; set; } = new();
        }
        """;

    [Test]
    [MethodDataSource(typeof(StructuralLockstepCases), nameof(StructuralLockstepCases.Names))]
    public async Task BuildReports(string name)
    {
        var structuralCase = StructuralLockstepCases.Get(name);
        var diagnostics = Run(GeneratorDriver.BuildDocxBytes(structuralCase.Paragraphs));

        await Assert.That(diagnostics.Select(_ => _.Id)).Contains(structuralCase.Id);
        await Assert.That(diagnostics.All(_ => _.Severity == DiagnosticSeverity.Error)).IsTrue();

        // PARCH005 has wording of its own; the other two carry the runtime's.
        if (structuralCase.Id == "PARCH005")
        {
            return;
        }

        var message = diagnostics.Single(_ => _.Id == structuralCase.Id).GetMessage();
        await Assert.That(message).Contains(structuralCase.Runtime);
    }

    [Test]
    public async Task UnparsableToken_NamesTheTokenAndTheParserError()
    {
        var diagnostics = Run(GeneratorDriver.BuildDocxBytes("{{ Name.ToUpper() }}"));

        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH006"]);
        var message = diagnostics.Single().GetMessage();
        await Assert.That(message).Contains("Model.parchment.docx");
        await Assert.That(message).Contains("[token: {{ Name.ToUpper() }}]");
    }

    [Test]
    public async Task UnbalancedTags_AreReportedOncePerPart()
    {
        var diagnostics = Run(
            GeneratorDriver.BuildDocxBytes(
                "{% endif %}",
                "{% for line in Lines %}",
                "{% endfor %}",
                "{% endfor %}"));

        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH035"]);
        await Assert.That(diagnostics.Single().GetMessage())
            .Contains("Unexpected '{% endif %}' without a matching opening tag.");
    }

    [Test]
    public async Task BalancedNesting_NoDiagnostic()
    {
        var diagnostics = Run(
            GeneratorDriver.BuildDocxBytes(
                "{% if Flag %}",
                "{% for line in Lines %}",
                "{% if line.Description %}",
                "{{ line.Description }}",
                "{% elsif Flag %}",
                "{{ Name }}",
                "{% else %}",
                "None",
                "{% endif %}",
                "{% endfor %}",
                "{% elsif Name %}",
                "{{ Name }}",
                "{% endif %}"));

        await Assert.That(diagnostics).IsEmpty();
    }

    // The runtime builds one scope tree per part, so a tag in the header cannot close a block
    // opened in the body, however the two parts are ordered in the package.
    [Test]
    public async Task BlockSplitBetweenHeaderAndBody_IsUnbalancedInBoth()
    {
        var diagnostics = Run(
            GeneratorDriver.BuildDocxBytesWithHeader(
                bodyParagraphs: ["{{ Name }}", "{% endif %}"],
                headerParagraphs: ["{% if Flag %}", "{{ Name }}"]));

        var messages = diagnostics
            .Where(_ => _.Id == "PARCH035")
            .Select(_ => _.GetMessage())
            .ToList();
        await Assert.That(messages.Count).IsEqualTo(2);
        await Assert.That(messages.Any(_ => _.Contains("Missing closing tag 'endif'."))).IsTrue();
        await Assert.That(messages.Any(_ => _.Contains("Unexpected '{% endif %}' without a matching opening tag."))).IsTrue();
    }

    [Test]
    public async Task BalancedBlocksInHeaderAndBody_NoDiagnostic()
    {
        var diagnostics = Run(
            GeneratorDriver.BuildDocxBytesWithHeader(
                bodyParagraphs: ["{% for line in Lines %}", "{{ line.Description }}", "{% endfor %}"],
                headerParagraphs: ["{% if Flag %}", "{{ Name }}", "{% endif %}"]));

        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task UnsupportedTag_LeavesBalanceUnreported()
    {
        // With a tag the generator cannot classify, which tag closes which block is a guess, so
        // PARCH003 stands alone.
        var diagnostics = Run(
            GeneratorDriver.BuildDocxBytes(
                "{% unless Flag %}",
                "{{ Name }}",
                "{% endif %}"));

        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH003"]);
    }

    static ImmutableArray<Diagnostic> Run(byte[] docx)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Model.parchment.docx", docx));
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult().Results.Single().Diagnostics;
    }
}
