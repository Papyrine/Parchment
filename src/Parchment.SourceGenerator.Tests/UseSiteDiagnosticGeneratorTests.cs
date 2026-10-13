/// <summary>
/// A member that is deprecated or experimental raises a compiler diagnostic wherever it is named,
/// and the generated accessors name every bindable member. These compile the generator's output
/// against the real assemblies, since the diagnostics come from the compiler rather than the
/// generator.
/// </summary>
public class UseSiteDiagnosticGeneratorTests
{
    static string Model(string members, string types = "", string attribute = "ParchmentModel") =>
        $$"""
          using System;
          using System.Collections.Generic;
          using System.Diagnostics.CodeAnalysis;
          using Parchment;

          namespace Sample;

          {{types}}

          [{{attribute}}]
          public partial class Legacy
          {
              public string Name { get; set; } = "";
              {{members}}
          }
          """;

    static GeneratorDriver.CompileResult Compile(string source, params string[] paragraphs) =>
        GeneratorDriver.Compile(
            source,
            new TemplateFile("Legacy.parchment.docx", GeneratorDriver.BuildDocxBytes(paragraphs.Length == 0 ? ["{{ Name }}"] : paragraphs)));

    static async Task AssertClean(GeneratorDriver.CompileResult result)
    {
        await Assert.That(result.Generator).IsEmpty();
        await Assert.That(result.Generated).IsEmpty();
    }

    static string? Pragma(GeneratorDriver.CompileResult result) =>
        result.Source
            .Split('\n')
            .Select(_ => _.TrimEnd('\r'))
            .FirstOrDefault(_ => _.StartsWith("#pragma warning disable", StringComparison.Ordinal));

    [Test]
    public async Task PlainModel_CompilesCleanWithoutAPragma()
    {
        var result = Compile(Model("public string? Note { get; set; }"));

        await AssertClean(result);
        await Assert.That(Pragma(result)).IsNull();
    }

    // Guards the harness: without the pragma the compiler does report the generated reference, so
    // an empty list in the other tests means suppressed rather than never looked at.
    [Test]
    public async Task ObsoleteMember_WarnsWhenThePragmaIsRemoved()
    {
        var source = Model("""[Obsolete("use Name")] public string OldName { get; set; } = "";""");
        var result = Compile(source);
        var tree = result.Output.SyntaxTrees.Single(GeneratorDriver.IsGenerated);
        var stripped = string.Join(
            '\n',
            tree.ToString().Split('\n').Where(_ => !_.StartsWith("#pragma", StringComparison.Ordinal)));

        var ids = result.Output
            .ReplaceSyntaxTree(tree, tree.WithChangedText(SourceText.From(stripped)))
            .GetDiagnostics()
            .Select(_ => _.Id)
            .ToList();

        await Assert.That(ids).Contains("CS0618");
    }

    [Test]
    public async Task ObsoleteMember_StaysBindable()
    {
        var result = Compile(
            Model(
                """
                [Obsolete("use Name")] public string OldName { get; set; } = "";
                [Obsolete] public string Older = "";
                [Obsolete("static")] public static string Oldest => "";
                """),
            "{{ OldName }} {{ Older }} {{ Oldest }}");

        await AssertClean(result);
        await Assert.That(Pragma(result)).IsEqualTo("#pragma warning disable CS0612, CS0618");
        await Assert.That(result.Source).Contains("new(\"OldName\"");
        await Assert.That(result.Source).Contains("new(\"Older\"");
        await Assert.That(result.Source).Contains("new(\"Oldest\"");
    }

    [Test]
    public async Task ObsoleteMember_Bindable()
    {
        var result = GeneratorDriver.Compile(
            Model(
                """[Obsolete("use Name")] public string OldName { get; set; } = "";""",
                attribute: "ParchmentBindable"));

        await AssertClean(result);
        await Assert.That(result.Source).Contains("new(\"OldName\"");
    }

    [Test]
    public Task ObsoleteMember_Markdown()
    {
        var result = GeneratorDriver.Compile(
            Model("""[Obsolete("use Name")] public string OldName { get; set; } = "";"""),
            new TemplateFile("Legacy.parchment.md", "{{ OldName }}"u8.ToArray()));

        return AssertClean(result);
    }

    [Test]
    public async Task ObsoleteMember_WithItsOwnDiagnosticId()
    {
        var result = Compile(
            Model("""[Obsolete("use Name", DiagnosticId = "LEG001")] public string OldName { get; set; } = "";"""),
            "{{ OldName }}");

        await AssertClean(result);
        await Assert.That(Pragma(result)).IsEqualTo("#pragma warning disable LEG001");
    }

    [Test]
    public Task ObsoleteMember_ThroughTheAttributeMaps()
    {
        var result = Compile(
            Model(
                """
                [Obsolete("old")] [Html] public string OldHtml { get; set; } = "";
                [Obsolete("old")] public List<string> OldTags { get; set; } = new();
                [Obsolete("old")] [EditableField] public string OldNote { get; set; } = "";
                [Obsolete("old")] [EditableField] public List<Row> OldRows { get; set; } = new();
                [Obsolete("old")] public Inner OldInner { get; set; } = new();
                """,
                """
                public class Row
                {
                    [Obsolete("old")] [EditableField] public string Cell { get; set; } = "";
                }

                public class Inner
                {
                    [Obsolete("old")] [Markdown] public string Body { get; set; } = "";
                }
                """),
            "{{ OldHtml }}",
            "{{ OldTags }}",
            "{{ OldNote }}",
            "{{ OldInner.Body }}");

        return AssertClean(result);
    }

    [Test]
    public async Task ObsoleteMemberType_StaysBindable()
    {
        var result = Compile(
            Model(
                """
                public Old Single { get; set; } = new();
                public List<Old> Many { get; set; } = new();
                public Old[] Array { get; set; } = [];
                public Outer.Nested Deep { get; set; } = new();
                """,
                """
                [Obsolete("legacy")]
                public class Old
                {
                    public string X { get; set; } = "";
                }

                [Obsolete("legacy")]
                public class Outer
                {
                    public class Nested
                    {
                        public string Y { get; set; } = "";
                    }
                }
                """),
            "{{ Single.X }} {{ Deep.Y }}",
            "{% for item in Many %}",
            "{{ item.X }}",
            "{% endfor %}");

        await AssertClean(result);
        await Assert.That(result.Source).Contains("typeof(global::Sample.Old)");
    }

    [Test]
    public Task ObsoleteModel_CompilesClean()
    {
        var result = Compile(
            """
            using System;
            using Parchment;

            namespace Sample;

            [Obsolete("use Letter")]
            [ParchmentModel]
            public partial class Legacy
            {
                public string Name { get; set; } = "";
            }
            """);

        return AssertClean(result);
    }

    [Test]
    public async Task ObsoleteAsErrorMember_IsNotBound()
    {
        var result = Compile(
            Model(
                """
                [Obsolete("gone", true)] public string Gone { get; set; } = "";
                public Removed Single { get; set; } = null!;
                public List<Removed> Many { get; set; } = null!;
                """,
                """
                [Obsolete("gone", true)]
                public class Removed
                {
                    public string X { get; set; } = "";
                }
                """));

        await AssertClean(result);
        await Assert.That(Pragma(result)).IsNull();
        await Assert.That(result.Source).DoesNotContain("Gone");
        await Assert.That(result.Source).DoesNotContain("Removed");
    }

    [Test]
    public async Task ObsoleteAsErrorMember_DoesNotExposeTheBaseMemberItHides()
    {
        var result = Compile(
            Model(
                """[Obsolete("gone", true)] public new string Title { get; set; } = "";""",
                """
                public class LegacyBase
                {
                    public string Title { get; set; } = "";
                }
                """)
                .Replace("partial class Legacy", "partial class Legacy : LegacyBase"));

        await AssertClean(result);
        await Assert.That(result.Source).DoesNotContain("Title");
    }

    [Test]
    public async Task ObsoleteAsErrorMember_TokenIsAnUnknownMember()
    {
        var result = Compile(
            Model("""[Obsolete("gone", true)] public string Gone { get; set; } = "";"""),
            "{{ Gone }}");

        await Assert.That(result.Generator.Select(_ => _.Id)).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task ExperimentalMember_StaysBindable()
    {
        var result = Compile(
            Model(
                """
                [Experimental("EXP001")] public string Beta { get; set; } = "";
                public Preview Next { get; set; } = null!;
                public List<Preview> Later { get; set; } = null!;
                """,
                """
                [Experimental("EXP002")]
                public class Preview
                {
                    public string X { get; set; } = "";
                }
                """),
            "{{ Beta }} {{ Next.X }}");

        // The model's own declarations of the experimental type are the consumer's to suppress.
        await Assert.That(result.Generator).IsEmpty();
        await Assert.That(result.Generated).IsEmpty();
        await Assert.That(Pragma(result)).IsEqualTo("#pragma warning disable EXP001, EXP002");
    }

    [Test]
    public async Task ExperimentalBaseClass_InheritedMemberStaysBindable()
    {
        var result = Compile(
            Model(
                "",
                """
                [Experimental("EXP003")]
                public class PreviewBase
                {
                    public string Early { get; set; } = "";
                }
                """)
                .Replace("partial class Legacy", "partial class Legacy : PreviewBase"),
            "{{ Early }}");

        await AssertClean(result);
        await Assert.That(Pragma(result)).IsEqualTo("#pragma warning disable EXP003");
    }

    // The trim and AOT analyzers are not part of this compilation, so only the emitted pragma can
    // be checked.
    [Test]
    public async Task RequiresAttributesOnAnAccessor_AreSuppressed()
    {
        var result = Compile(
            Model(
                """
                public string Reflected
                {
                    [RequiresUnreferencedCode("reflects")]
                    [RequiresDynamicCode("emits")]
                    get => "";
                }
                """));

        await AssertClean(result);
        await Assert.That(Pragma(result)).IsEqualTo("#pragma warning disable IL2026, IL3050");
    }
}
