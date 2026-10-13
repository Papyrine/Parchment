/// <summary>
/// A model is free to use any legal C# name, so every identifier the generator writes has to be
/// written the way the language requires it, and every hint name has to be one Roslyn accepts and
/// no other target shares. The same shapes are compiled for real by
/// <c>Parchment.Tests/KeywordIdentifierTests</c>.
/// </summary>
public class GeneratedIdentifierTests
{
    static GeneratorRunResult Run(string source, params TemplateFile[] files)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(source, files);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult().Results.Single();
    }

    static TemplateFile Docx(string typeName, params string[] paragraphs) =>
        new($"{typeName}.parchment.docx", GeneratorDriver.BuildDocxBytes(paragraphs));

    static async Task<string> AssertGenerated(GeneratorRunResult result, params string[] hintNames)
    {
        await Assert.That(result.Exception).IsNull();
        await Assert.That(result.Diagnostics.Select(_ => _.Id)).IsEmpty();
        var syntaxErrors = result.GeneratedSources
            .SelectMany(_ => _.SyntaxTree.GetDiagnostics())
            .Select(_ => _.ToString())
            .ToList();
        await Assert.That(syntaxErrors).IsEmpty();
        await Assert.That(result.GeneratedSources.Select(_ => _.HintName).Order(StringComparer.Ordinal))
            .IsEquivalentTo(hintNames.Order(StringComparer.Ordinal));
        return string.Concat(result.GeneratedSources.Select(_ => _.SourceText.ToString()));
    }

    static string Source(GeneratorRunResult result, string hintName) =>
        result.GeneratedSources.Single(_ => _.HintName == hintName).SourceText.ToString();

    [Test]
    public async Task KeywordMembers_AreEscapedInCodeAndNotInTemplateNames()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public sealed class HtmlAttribute : System.Attribute;

            [ParchmentModel]
            public partial class Letter
            {
                public string @class { get; set; } = "";
                public string @event = "";
                public static string @default => "";
                public Inner @object { get; set; } = new();
                public List<string> @params { get; set; } = new();
                [EditableField] public string @string { get; set; } = "";
                [ExcelsiorTable(Configure = nameof(@static))] public List<Inner> @foreach { get; set; } = new();
                static void @static(object builder) { }
            }

            public class Inner
            {
                [Html] public string @namespace { get; set; } = "";
                [EditableField] public string @new { get; set; } = "";
            }
            """;

        var result = Run(source, Docx("Letter", "{{ class }} {{ event }} {{ object.namespace }}", "{{ string }}", "{{ object.new }}", "{{ foreach }}"));
        var generated = await AssertGenerated(result, "Sample_Letter_ParchmentModel.g.cs");

        await Assert.That(generated).Contains("new(\"class\", new global::Fluid.Accessors.DelegateAccessor((o, _) => ((global::Sample.Letter)o).@class))");
        await Assert.That(generated).Contains("new(\"event\", new global::Fluid.Accessors.DelegateAccessor((o, _) => ((global::Sample.Letter)o).@event))");
        await Assert.That(generated).Contains("(_, _) => global::Sample.Letter.@default");
        await Assert.That(generated).Contains("new(\"object.namespace\", global::Parchment.Generated.FormatMapKind.Html, o => ((global::Sample.Letter)o).@object?.@namespace)");
        await Assert.That(generated).Contains("new(\"params\", o => ((global::Sample.Letter)o).@params)");
        await Assert.That(generated).Contains("((global::Sample.Letter)o).@string = (string)v!;");
        await Assert.That(generated).Contains("var p = ((global::Sample.Letter)o).@object; if (p != null) { p.@new = (string)v!; }");
        await Assert.That(generated).Contains("o => ((global::Sample.Letter)o).@object != null");
        await Assert.That(generated).Contains("b => global::Sample.Letter.@static((global::Excelsior.WordTableBuilder<global::Sample.Inner>)b)");
    }

    [Test]
    public async Task KeywordModelType_IsEscapedInThePartialDeclaration()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public partial class @class
            {
                [ParchmentModel]
                public partial class @event
                {
                    public string Name { get; set; } = "";
                }
            }
            """;

        var result = Run(source, Docx("event", "{{ Name }}"));
        var generated = await AssertGenerated(result, "Sample_class_event_ParchmentModel.g.cs");

        await Assert.That(generated).Contains("partial class @class");
        await Assert.That(generated).Contains("partial class @event");
        await Assert.That(generated).Contains("typeof(global::Sample.@class.@event)");
    }

    [Test]
    public async Task KeywordNamespaceAndMemberType()
    {
        var source =
            """
            using Parchment;

            namespace Sample.@namespace;

            public class @class
            {
                public string Name { get; set; } = "";
            }

            [ParchmentModel]
            public partial class Letter
            {
                public @class Thing { get; set; } = new();
            }
            """;

        var result = Run(source, Docx("Letter", "{{ Thing.Name }}"));
        var generated = await AssertGenerated(result, "Sample_namespace_Letter_ParchmentModel.g.cs");

        await Assert.That(generated).Contains("namespace Sample.@namespace;");
        await Assert.That(generated).Contains("typeof(global::Sample.@namespace.@class)");
    }

    // `file`, `required` and `scoped` cannot name a type unless escaped; `record` and `var` only warn.
    [Test]
    [Arguments("record")]
    [Arguments("var")]
    [Arguments("async")]
    [Arguments("file")]
    [Arguments("required")]
    [Arguments("scoped")]
    public async Task ContextualKeywordNames(string name)
    {
        var source =
            $$"""
              using Parchment;

              namespace Sample.@{{name}};

              [ParchmentBindable]
              public partial class @{{name}}
              {
                  [EditableField] public string @{{name}} { get; set; } = "";
              }
              """;

        var result = Run(source);
        var generated = await AssertGenerated(result, $"Sample_{name}_{name}_ParchmentBindable.g.cs");

        await Assert.That(generated).Contains($"partial class @{name}");
        await Assert.That(generated).Contains($"new(\"{name}\", ");

        // What the escaping is for: with the generated tree in, the declaration is still legal.
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        setup.Driver.RunGeneratorsAndUpdateCompilation(setup.Compilation, out var output, out _);
        var declarationErrors = output.GetDeclarationDiagnostics()
            .Where(_ => _.Severity == DiagnosticSeverity.Error)
            .Where(_ => _.Id is "CS9056" or "CS9029" or "CS9062" or "CS0102" or "CS0260")
            .Select(_ => _.ToString());
        await Assert.That(declarationErrors).IsEmpty();
    }

    const string caseTwins =
        """
        namespace Sample
        {
            [Parchment.ParchmentBindable]
            public partial class Letter
            {
                public string Upper { get; set; } = "";
            }
        }
        """;

    const string caseTwinsLower =
        """
        namespace Sample
        {
            [Parchment.ParchmentBindable]
            public partial class letter
            {
                public string Lower { get; set; } = "";
            }
        }
        """;

    // Hint names become file names, so two that differ only by case are one name to Roslyn.
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task TypesDifferingOnlyByCase_BothGenerate(bool upperFirst)
    {
        var source = upperFirst ? caseTwins + caseTwinsLower : caseTwinsLower + caseTwins;

        var result = Run(source);
        await AssertGenerated(result, "Sample_Letter_ParchmentBindable.g.cs", "Sample_letter_2_ParchmentBindable.g.cs");

        // The same type gets the same name whichever order the declarations arrive in.
        await Assert.That(Source(result, "Sample_Letter_ParchmentBindable.g.cs")).Contains("\"Upper\"");
        await Assert.That(Source(result, "Sample_letter_2_ParchmentBindable.g.cs")).Contains("\"Lower\"");

        // Declared bare, an all-lower-case type name is CS8981 in the generated file.
        await Assert.That(Source(result, "Sample_letter_2_ParchmentBindable.g.cs")).Contains("partial class @letter");
    }

    [Test]
    public async Task NamesThatFlattenTheSame_BothGenerate_Bindable()
    {
        var source =
            """
            namespace A_B
            {
                [Parchment.ParchmentBindable]
                public partial class C
                {
                    public string First { get; set; } = "";
                }
            }

            namespace A
            {
                [Parchment.ParchmentBindable]
                public partial class B_C
                {
                    public string Second { get; set; } = "";
                }
            }

            // Already owns the name the first collision would otherwise be given.
            namespace A_B
            {
                [Parchment.ParchmentBindable]
                public partial class C_2
                {
                    public string Third { get; set; } = "";
                }
            }
            """;

        var result = Run(source);
        await AssertGenerated(
            result,
            "A_B_C_ParchmentBindable.g.cs",
            "A_B_C_3_ParchmentBindable.g.cs",
            "A_B_C_2_ParchmentBindable.g.cs");

        await Assert.That(Source(result, "A_B_C_ParchmentBindable.g.cs")).Contains("\"Second\"");
        await Assert.That(Source(result, "A_B_C_3_ParchmentBindable.g.cs")).Contains("\"First\"");
        await Assert.That(Source(result, "A_B_C_2_ParchmentBindable.g.cs")).Contains("\"Third\"");
    }

    [Test]
    public async Task NamesThatFlattenTheSame_BothGenerate_Model()
    {
        var source =
            """
            namespace A_B
            {
                [Parchment.ParchmentModel]
                public partial class C
                {
                    public string First { get; set; } = "";
                }
            }

            namespace A
            {
                [Parchment.ParchmentModel]
                public partial class B_C
                {
                    public string Second { get; set; } = "";
                }
            }
            """;

        var result = Run(source, Docx("C", "{{ First }}"), Docx("B_C", "{{ Second }}"));
        await AssertGenerated(result, "A_B_C_ParchmentModel.g.cs", "A_B_C_2_ParchmentModel.g.cs");

        await Assert.That(Source(result, "A_B_C_ParchmentModel.g.cs")).Contains("\"Second\"");
        await Assert.That(Source(result, "A_B_C_2_ParchmentModel.g.cs")).Contains("\"First\"");
    }

    // The accessor tables are fields named after the type they describe, and two types of one
    // model can flatten to the same field name.
    [Test]
    public async Task MemberTypesThatFlattenTheSame_GetDistinctFields()
    {
        var source =
            """
            namespace A_B
            {
                public class C
                {
                    public string First { get; set; } = "";
                }
            }

            namespace A
            {
                public class B_C
                {
                    public string Second { get; set; } = "";
                }
            }

            namespace Sample
            {
                [Parchment.ParchmentBindable]
                public partial class Letter
                {
                    public A_B.C One { get; set; } = new();
                    public A.B_C Two { get; set; } = new();
                }
            }
            """;

        var result = Run(source);
        var generated = await AssertGenerated(result, "Sample_Letter_ParchmentBindable.g.cs");

        await Assert.That(generated).Contains("RegisterFluidAccessors(typeof(global::A_B.C), _Accessors_global__A_95_B_C);");
        await Assert.That(generated).Contains("RegisterFluidAccessors(typeof(global::A.B_C), _Accessors_global__A_B_95_C);");
    }
}
