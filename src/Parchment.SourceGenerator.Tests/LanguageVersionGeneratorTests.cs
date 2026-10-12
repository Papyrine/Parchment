using System.Globalization;

/// <summary>
/// The generated file is compiled with the consumer's <c>LangVersion</c>, so every construct it
/// uses is a requirement on that project. C# 9 is the floor — a module initializer cannot be
/// written in anything older — and nothing emitted may ask for more than that.
/// </summary>
public class LanguageVersionGeneratorTests
{
    // One member of every kind that has its own emission: plain accessors, an enum, a nested
    // path, each editable kind including a collection, both format markers, a string list and an
    // Excelsior table with a Configure call.
    const string richModel =
        """
        using System;
        using System.Collections.Generic;
        using Parchment;

        namespace Sample
        {
            public enum Status
            {
                Draft,
                Final
            }

            public class Customer
            {
                [EditableField]
                public string? Email { get; set; }
            }

            public class Line
            {
                public string Description { get; set; } = "";
            }

            public class Row
            {
                [EditableField]
                public string Note { get; set; } = "";
            }

            public class HtmlAttribute : Attribute
            {
            }

            public class MarkdownAttribute : Attribute
            {
            }

            [ParchmentModel(Protection = ProtectionMode.None)]
            public partial class Order
            {
                public string Number { get; set; } = "";

                [EditableField]
                public bool Approved { get; set; }

                [EditableField(DateFormat = "dd MMM yyyy")]
                public DateTime Delivery { get; set; }

                [EditableField]
                public Status Status { get; set; }

                [EditableField]
                public decimal? Discount { get; set; }

                [EditableField(MultiLine = true)]
                public string Instructions { get; set; } = "";

                [EditableField]
                [Html]
                public string Terms { get; set; } = "";

                [EditableField]
                public List<Row> Rows { get; set; } = new List<Row>();

                [Markdown]
                public string Notes { get; set; } = "";

                public List<string> Tags { get; set; } = new List<string>();

                public Customer Customer { get; set; } = new Customer();

                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new List<Line>();

                internal static void ConfigureLines(object builder)
                {
                }
            }
        }
        """;

    static readonly string[] richTemplate =
    [
        "{{ Number }}",
        "{{ Approved }}",
        "{{ Delivery }}",
        "{{ Status }}",
        "{{ Discount }}",
        "{{ Instructions }}",
        "{{ Terms }}",
        "{% for row in Rows %}",
        "{{ row.Note }}",
        "{% endfor %}",
        "{{ Notes }}",
        "{{ Tags }}",
        "{{ Customer.Email }}",
        "{{ Lines }}"
    ];

    [Test]
    public async Task CSharp9_DocxModel_UsesNothingNewer()
    {
        var output = Run(
            richModel,
            LanguageVersion.CSharp9,
            new TemplateFile("Order.parchment.docx", GeneratorDriver.BuildDocxBytes(richTemplate)));

        await Assert.That(output.GeneratorDiagnostics).IsEmpty();
        await Assert.That(output.Sources.Count).IsEqualTo(1);
        await Assert.That(output.LanguageErrors).IsEmpty();
        await Assert.That(output.Sources[0]).Contains("namespace Sample\n{\n");
        await Assert.That(output.Sources[0]).EndsWith("}\n}\n");
    }

    [Test]
    public async Task CSharp9_MarkdownModel_UsesNothingNewer()
    {
        var output = Run(
            richModel,
            LanguageVersion.CSharp9,
            new TemplateFile("Order.parchment.md", "# {{ Number }}\n\n{{ Customer.Email }}\n"u8.ToArray()));

        await Assert.That(output.GeneratorDiagnostics).IsEmpty();
        await Assert.That(output.Sources.Count).IsEqualTo(1);
        await Assert.That(output.LanguageErrors).IsEmpty();
    }

    // The namespace is the one construct that differs, and it is absent, nested under partial
    // types, or shared with a bindable depending on where the model sits.
    [Test]
    public async Task CSharp9_EveryPlacement_UsesNothingNewer()
    {
        var source =
            """
            using Parchment;

            [ParchmentModel]
            public partial class Top
            {
                public string Name { get; set; } = "";
            }

            namespace Sample.Deep
            {
                public partial class Outer
                {
                    public partial record Inner
                    {
                        [ParchmentModel]
                        public partial class Nested
                        {
                            public string Name { get; set; } = "";
                        }
                    }
                }

                [ParchmentModel]
                public partial record Positional(string Name);

                [ParchmentBindable]
                public partial class Bound
                {
                    public string Name { get; set; } = "";
                }
            }
            """;
        var template = GeneratorDriver.BuildDocxBytes("{{ Name }}");
        var output = Run(
            source,
            LanguageVersion.CSharp9,
            new TemplateFile("Top.parchment.docx", template),
            new TemplateFile("Nested.parchment.docx", template),
            new TemplateFile("Positional.parchment.docx", template));

        await Assert.That(output.GeneratorDiagnostics).IsEmpty();
        await Assert.That(output.Sources.Count).IsEqualTo(4);
        await Assert.That(output.LanguageErrors).IsEmpty();
        await Assert.That(output.Sources.Count(_ => _.Contains("namespace Sample.Deep\n{\n"))).IsEqualTo(3);
        await Assert.That(output.Sources.Any(_ => _.Contains("namespace Sample.Deep;"))).IsFalse();
    }

    // From C# 10 on the file-scoped form stays, so the output of every project that does not pin
    // an old LangVersion is unchanged.
    [Test]
    [Arguments(LanguageVersion.CSharp10)]
    [Arguments(LanguageVersion.CSharp12)]
    [Arguments(LanguageVersion.Latest)]
    [Arguments(LanguageVersion.Default)]
    [Arguments(LanguageVersion.Preview)]
    public async Task CSharp10AndLater_KeepsFileScopedNamespace(LanguageVersion version)
    {
        var output = Run(
            richModel,
            version,
            new TemplateFile("Order.parchment.docx", GeneratorDriver.BuildDocxBytes(richTemplate)));

        await Assert.That(output.GeneratorDiagnostics).IsEmpty();
        await Assert.That(output.LanguageErrors).IsEmpty();
        await Assert.That(output.Sources[0]).Contains("namespace Sample;\n");
    }

    // Guards the check itself: a C# 8 project cannot take a module initializer at all, and the
    // compiler says so in the terms LanguageErrors looks for.
    [Test]
    public async Task CSharp8_IsBelowTheFloor()
    {
        var output = Run(
            richModel,
            LanguageVersion.CSharp8,
            new TemplateFile("Order.parchment.docx", GeneratorDriver.BuildDocxBytes(richTemplate)));

        await Assert.That(output.LanguageErrors).IsNotEmpty();
    }

    sealed record Output(
        IReadOnlyList<string> GeneratorDiagnostics,
        IReadOnlyList<string> Sources,
        IReadOnlyList<string> LanguageErrors);

    static Output Run(string source, LanguageVersion version, params TemplateFile[] files)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            files,
            modelFileName: null,
            languageVersion: version);
        var result = setup.Driver
            .RunGeneratorsAndUpdateCompilation(setup.Compilation, out var compilation, out _)
            .GetRunResult();

        var generated = result.GeneratedTrees.ToHashSet();
        // Syntax errors, and the "feature is not available in C# n" family, which the compiler
        // words the same way whichever feature it is. The stub compilation references neither
        // Parchment nor Fluid, so unresolved-name errors are expected and not what is asked here.
        var languageErrors = compilation.GetDiagnostics()
            .Where(_ => _.Location.SourceTree != null && generated.Contains(_.Location.SourceTree))
            .Where(_ => _.Severity == DiagnosticSeverity.Error)
            .Where(_ => generated.Any(tree => tree.GetDiagnostics().Contains(_)) ||
                        _.GetMessage(CultureInfo.InvariantCulture).Contains("language version"))
            .Select(_ => $"{_.Id}: {_.GetMessage(CultureInfo.InvariantCulture)}")
            .Distinct()
            .ToList();

        return new(
            result.Diagnostics.Select(_ => $"{_.Id}: {_.GetMessage(CultureInfo.InvariantCulture)}").ToList(),
            result.GeneratedTrees.Select(_ => _.ToString().Replace("\r\n", "\n")).ToList(),
            languageErrors);
    }
}
