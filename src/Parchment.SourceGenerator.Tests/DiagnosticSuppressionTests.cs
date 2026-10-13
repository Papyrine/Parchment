// Every diagnostic is reported at the model's attribute. These pin that the location is a real
// source location, the only kind the compiler matches against an in-source suppression.
public class DiagnosticSuppressionTests
{
    const string htmlAttribute =
        """

        namespace Parchment
        {
            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class HtmlAttribute : System.Attribute;
        }
        """;

    const string model =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Letter
        {
            [Html]
            public static string Banner { get; set; } = "";

            public string Name { get; set; } = "";
        }
        """;

    [Test]
    public async Task Unsuppressed_ReportedInSource()
    {
        var diagnostics = Run(model);

        await Assert.That(Ids(diagnostics)).IsEquivalentTo(["PARCH012", "PARCH019"]);
        foreach (var diagnostic in diagnostics)
        {
            await Assert.That(diagnostic.Location.IsInSource).IsTrue();
            await Assert.That((await diagnostic.Location.SourceTree!.GetTextAsync()).ToString(diagnostic.Location.SourceSpan))
                .IsEqualTo("ParchmentModel");
        }
    }

    [Test]
    public async Task Pragma_Suppresses()
    {
        var source = model
            .Replace("[ParchmentModel]", "#pragma warning disable PARCH012, PARCH019\n[ParchmentModel]\n#pragma warning restore PARCH012, PARCH019");

        await Assert.That(Ids(Run(source))).IsEmpty();
    }

    [Test]
    public async Task Pragma_SuppressesOnlyTheNamedId()
    {
        var source = model
            .Replace("[ParchmentModel]", "#pragma warning disable PARCH019\n[ParchmentModel]\n#pragma warning restore PARCH019");

        await Assert.That(Ids(Run(source))).IsEquivalentTo(["PARCH012"]);
    }

    [Test]
    public async Task SuppressMessageAttribute_Suppresses()
    {
        var source = model
            .Replace(
                "[ParchmentModel]",
                """
                [ParchmentModel]
                [System.Diagnostics.CodeAnalysis.SuppressMessage("Parchment", "PARCH012")]
                [System.Diagnostics.CodeAnalysis.SuppressMessage("Parchment", "PARCH019")]
                """);

        await Assert.That(Ids(Run(source))).IsEmpty();
    }

    [Test]
    public async Task PerFileSeverity_Applies()
    {
        var diagnostics = Run(
            model,
            new()
            {
                ["PARCH012"] = ReportDiagnostic.Suppress,
                ["PARCH019"] = ReportDiagnostic.Error
            });

        await Assert.That(Ids(diagnostics)).IsEquivalentTo(["PARCH019"]);
        await Assert.That(diagnostics.Single().Severity).IsEqualTo(DiagnosticSeverity.Error);
    }

    [Test]
    public async Task Bindable_PragmaSuppresses()
    {
        var bindable = model.Replace("[ParchmentModel]", "[ParchmentBindable]");
        await Assert.That(Ids(Run(bindable))).IsEquivalentTo(["PARCH019"]);

        var source = bindable
            .Replace("[ParchmentBindable]", "#pragma warning disable PARCH019\n[ParchmentBindable]\n#pragma warning restore PARCH019");
        await Assert.That(Ids(Run(source))).IsEmpty();
    }

    [Test]
    public async Task Markdown_ReportedInSource()
    {
        // The error is one no suppression applies to, so the location
        // kind is asserted directly.
        var template = new TemplateFile("Letter.parchment.md", "{{ Missing }}\n"u8.ToArray());
        var diagnostic = Run(model, template: template).Single(_ => _.Id == "PARCH001");

        await Assert.That(diagnostic.Id).IsEqualTo("PARCH001");
        await Assert.That(diagnostic.Location.IsInSource).IsTrue();
    }

    static List<string> Ids(List<Diagnostic> diagnostics) =>
        diagnostics
            .Select(_ => _.Id)
            .OrderBy(_ => _, StringComparer.Ordinal)
            .ToList();

    // The diagnostics a build would show: the driver drops what the compilation options turn off
    // and marks what a pragma or an attribute suppresses.
    static List<Diagnostic> Run(
        string source,
        Dictionary<string, ReportDiagnostic>? perFileSeverity = null,
        TemplateFile? template = null)
    {
        template ??= new("Letter.parchment.docx", GeneratorDriver.BuildDocxBytesWithoutPrivacyFlag("{{ Name }}"));
        var setup = GeneratorDriver.CreateDriverWithDocxes(source + htmlAttribute, template);
        var compilation = setup.Compilation;
        if (perFileSeverity != null)
        {
            compilation = compilation.WithOptions(
                compilation.Options.WithSyntaxTreeOptionsProvider(new SeverityProvider(compilation.SyntaxTrees[1], perFileSeverity)));
        }

        return setup.Driver
            .RunGenerators(compilation)
            .GetRunResult()
            .Diagnostics
            .Where(_ => !_.IsSuppressed)
            .ToList();
    }

    // Stands in for an .editorconfig section that matches the model's file only.
    sealed class SeverityProvider(SyntaxTree modelTree, Dictionary<string, ReportDiagnostic> severities) :
        SyntaxTreeOptionsProvider
    {
        public override GeneratedKind IsGenerated(SyntaxTree tree, Cancel cancel) =>
            GeneratedKind.NotGenerated;

        public override bool TryGetDiagnosticValue(SyntaxTree tree, string diagnosticId, Cancel cancel, out ReportDiagnostic severity)
        {
            if (tree == modelTree)
            {
                return severities.TryGetValue(diagnosticId, out severity);
            }

            severity = ReportDiagnostic.Default;
            return false;
        }

        public override bool TryGetGlobalDiagnosticValue(string diagnosticId, Cancel cancel, out ReportDiagnostic severity)
        {
            severity = ReportDiagnostic.Default;
            return false;
        }
    }
}
