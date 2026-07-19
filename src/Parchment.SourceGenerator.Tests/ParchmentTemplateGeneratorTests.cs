public class ParchmentTemplateGeneratorTests
{
    const string letterModelDocx =
        """
        using Parchment;

        namespace Sample;

        public class Customer
        {
            public string Name { get; set; } = "";
        }

        [ParchmentModel("template.docx")]
        public partial class Letter
        {
            public Customer Customer { get; set; } = new();
        }
        """;

    const string letterModelMd =
        """
        using Parchment;

        namespace Sample;

        public class Customer
        {
            public string Name { get; set; } = "";
        }

        [ParchmentModel("template.md")]
        public partial class Letter
        {
            public Customer Customer { get; set; } = new();
        }
        """;

    // Same shape as the inline source in Markdown_ForLoop_Valid, shared by the liquid-identifier
    // tests below so the allowlist cases read as a set.
    const string invoiceModelMd =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        [ParchmentModel("template.md")]
        public partial class Invoice
        {
            public List<Line> Lines { get; set; } = new();
        }
        """;

    [Test]
    public Task Substitution_Valid()
    {
        var result = GeneratorDriver.Run(letterModelDocx, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public Task Substitution_MissingMember()
    {
        var result = GeneratorDriver.Run(letterModelDocx, "{{ Customer.Missing }}");
        return Verify(result);
    }

    [Test]
    public Task Substitution_IndexerStringLiteral_Valid()
    {
        // `{{ Customer['Name'] }}` is Fluid-equivalent to `{{ Customer.Name }}` — validation must
        // walk the indexer segment as if it were dotted.
        var result = GeneratorDriver.Run(letterModelDocx, "Hello {{ Customer['Name'] }}!");
        return Verify(result);
    }

    [Test]
    public Task Substitution_IndexerStringLiteral_MissingMember()
    {
        var result = GeneratorDriver.Run(letterModelDocx, "{{ Customer['Missing'] }}");
        return Verify(result);
    }

    [Test]
    public Task ForLoop_Valid()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                public string Description { get; set; } = "";
            }

            [ParchmentModel("template.docx")]
            public partial class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "{{ line.Description }}",
            "{% endfor %}");
        return Verify(result);
    }

    [Test]
    public async Task ForLoop_NestedShadowedVariable_OuterRestoredAfterInner()
    {
        // Outer loop binds `line` to OuterLine. Inner loop shadows with same name `line` bound to
        // InnerLine. After the inner `{% endfor %}`, the outer's `{{ line.OuterDescription }}` must
        // still resolve against OuterLine. Bug was: ForClose did `scope.Remove("line")` which
        // wiped the outer's binding instead of restoring it.
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class OuterLine
            {
                public string OuterDescription { get; set; } = "";
                public List<InnerLine> Inners { get; set; } = new();
            }

            public class InnerLine
            {
                public string InnerDescription { get; set; } = "";
            }

            [ParchmentModel("template.docx")]
            public partial class Doc
            {
                public List<OuterLine> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "outer={{ line.OuterDescription }}",
            "{% for line in line.Inners %}",
            "inner={{ line.InnerDescription }}",
            "{% endfor %}",
            "after-inner={{ line.OuterDescription }}",
            "{% endfor %}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public Task ForLoop_SourceNotEnumerable()
    {
        var result = GeneratorDriver.Run(
            letterModelDocx,
            "{% for item in Customer %}",
            "x",
            "{% endfor %}");
        return Verify(result);
    }

    [Test]
    public Task UnknownBlockTag()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Empty;
            """;
        var result = GeneratorDriver.Run(source, "{% foobar %}");
        return Verify(result);
    }

    [Test]
    public Task TemplateFileMissing()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("does-not-exist.docx")]
            public partial class Empty;
            """;
        var result = GeneratorDriver.Run(source, "ignored");
        return Verify(result);
    }

    [Test]
    public Task ForLoop_StringIsNotEnumerable()
    {
        // Guards against string being silently treated as IEnumerable<char>. ShapeBuilder
        // has a special-case for SpecialType.System_String; this test fails loudly if that
        // guard is removed (loop would validate as legal "for c in <string>" instead).
        var result = GeneratorDriver.Run(
            letterModelDocx,
            "{% for c in Customer.Name %}",
            "x",
            "{% endfor %}");
        return Verify(result);
    }

    [Test]
    public Task MultiTarget_MultiDocx()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            [ParchmentModel("letter.docx")]
            public partial class Letter
            {
                public Customer Customer { get; set; } = new();
            }

            [ParchmentModel("invoice.docx")]
            public partial class Invoice
            {
                public decimal Total { get; set; }
            }
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            ("letter.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!")),
            ("invoice.docx", GeneratorDriver.BuildDocxBytes("Total: {{ Total }}")));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return Verify(result);
    }

    [Test]
    public async Task TemplateReadError_CorruptDocx()
    {
        // PARCH006: exception message is platform/culture dependent, so assert the diagnostic
        // id directly instead of snapshotting the full message.
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Empty;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            ("template.docx", "not a zip file"u8.ToArray()));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var diagnostics = result.Results.Single().Diagnostics;

        await Assert.That(diagnostics.Length).IsEqualTo(1);
        await Assert.That(diagnostics[0].Id).IsEqualTo("PARCH006");
    }

    [Test]
    public async Task RemovePersonalInformation_Missing_EmitsWarning()
    {
        // PARCH012 fires when the template's word/settings.xml lacks the
        // <w:removePersonalInformation/> element, which corresponds to Word's "Remove personal
        // information from file properties on save" privacy option.
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Empty;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            ("template.docx", GeneratorDriver.BuildDocxBytesWithoutPrivacyFlag("hello")));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH012")).IsTrue();
        await Assert.That(diagnostics.Single(_ => _.Id == "PARCH012").Severity).IsEqualTo(DiagnosticSeverity.Warning);
    }

    [Test]
    public async Task RemovePersonalInformation_Present_NoWarning()
    {
        // Sanity check that BuildDocxBytes (which sets the flag) doesn't trip PARCH012.
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Empty;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            ("template.docx", GeneratorDriver.BuildDocxBytes("hello")));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH012")).IsFalse();
    }

    const string excelsiorModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        [ParchmentModel("template.docx")]
        public partial class Invoice
        {
            [ExcelsiorTable]
            public List<Line> Lines { get; set; } = new();
        }
        """;

    [Test]
    public async Task ExcelsiorToken_MixedInline_Diagnostic()
    {
        // Token in the same paragraph as other text must trip PARCH007. The SG asserts diagnostic
        // id directly instead of snapshotting because diagnostic messages may reword over time.
        var result = GeneratorDriver.Run(excelsiorModel, "Prefix {{ Lines }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsTrue();
    }

    [Test]
    public async Task ExcelsiorToken_WithFilter_Diagnostic()
    {
        // Filter chain on an [ExcelsiorTable] substitution must trip PARCH008.
        var result = GeneratorDriver.Run(excelsiorModel, "{{ Lines | reverse }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH008")).IsTrue();
    }

    [Test]
    public async Task ExcelsiorToken_Clean_NoDiagnostics()
    {
        // Baseline: plain {{ Lines }} in its own paragraph should NOT trip PARCH007/PARCH008.
        var result = GeneratorDriver.Run(excelsiorModel, "{{ Lines }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH008")).IsFalse();
    }

    const string formatModel =
        """
        using System.Diagnostics.CodeAnalysis;
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute { }

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class MarkdownAttribute : System.Attribute { }

        [ParchmentModel("template.docx")]
        public partial class Doc
        {
            [Html]
            public string Body { get; set; } = "";

            [Markdown]
            public string Notes { get; set; } = "";

            [StringSyntax("html")]
            public string Summary { get; set; } = "";
        }
        """;

    [Test]
    public async Task FormatToken_MixedInline_NoDiagnostic()
    {
        // Non-solo `[Html]`/`[Markdown]` tokens are allowed — the runtime splices inline content
        // into the host paragraph and splits the host paragraph for block-level content. PARCH009
        // (the legacy "must sit alone" diagnostic) is no longer emitted.
        var result = GeneratorDriver.Run(formatModel, "Prefix {{ Body }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH009")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsFalse();
    }

    [Test]
    public async Task FormatToken_WithFilter_Diagnostic()
    {
        var result = GeneratorDriver.Run(formatModel, "{{ Body | upcase }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsTrue();
    }

    [Test]
    public async Task FormatToken_MarkdownClean_NoDiagnostics()
    {
        var result = GeneratorDriver.Run(formatModel, "{{ Notes }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsFalse();
    }

    [Test]
    public async Task FormatToken_StringSyntaxHtml_NoDiagnostics()
    {
        var result = GeneratorDriver.Run(formatModel, "{{ Summary }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH009")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsFalse();
    }

    [Test]
    public Task Markdown_Substitution_Valid()
    {
        var result = GeneratorDriver.RunMarkdown(letterModelMd, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public Task Markdown_Substitution_MissingMember()
    {
        var result = GeneratorDriver.RunMarkdown(letterModelMd, "{{ Customer.Missing }}");
        return Verify(result);
    }

    [Test]
    public Task Markdown_Substitution_IndexerStringLiteral_MissingMember()
    {
        var result = GeneratorDriver.RunMarkdown(letterModelMd, "{{ Customer['Missing'] }}");
        return Verify(result);
    }

    [Test]
    public Task Markdown_ForLoop_Valid()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                public string Description { get; set; } = "";
            }

            [ParchmentModel("template.md")]
            public partial class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.RunMarkdown(
            source,
            """
            {% for line in Lines %}
            - {{ line.Description }}
            {% endfor %}
            """);
        return Verify(result);
    }

    [Test]
    public Task Markdown_ForLoop_MissingMemberInBody()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                public string Description { get; set; } = "";
            }

            [ParchmentModel("template.md")]
            public partial class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.RunMarkdown(
            source,
            """
            {% for line in Lines %}
            - {{ line.Missing }}
            {% endfor %}
            """);
        return Verify(result);
    }

    [Test]
    public Task Markdown_ForLoop_SourceNotEnumerable()
    {
        var result = GeneratorDriver.RunMarkdown(
            letterModelMd,
            """
            {% for item in Customer %}
            x
            {% endfor %}
            """);
        return Verify(result);
    }

    // The generator has to accept every liquid-introduced identifier the runtime markdown validator
    // accepts. It emits PARCH001 at Error, so a disagreement here fails the build for a template
    // that registers and renders correctly.
    [Test]
    public Task Markdown_ForLoopIdentifier_Accepted()
    {
        var result = GeneratorDriver.RunMarkdown(
            invoiceModelMd,
            """
            {% for line in Lines %}
            {{ forloop.index }}. {{ line.Description }}
            {% endfor %}
            """);
        return Verify(result);
    }

    [Test]
    public Task Markdown_AssignTarget_Accepted()
    {
        var result = GeneratorDriver.RunMarkdown(
            invoiceModelMd,
            """
            {% assign total = Lines %}
            {% for line in total %}
            - {{ line.Description }}
            {% endfor %}
            """);
        return Verify(result);
    }

    [Test]
    public Task Markdown_CaptureTarget_Accepted()
    {
        var result = GeneratorDriver.RunMarkdown(
            letterModelMd,
            "{% capture heading %}Report{% endcapture %}{{ heading }}");
        return Verify(result);
    }

    // The target is untyped, but the value is a normal expression and stays checked.
    [Test]
    public Task Markdown_AssignOfUnknownMember_Reported()
    {
        var result = GeneratorDriver.RunMarkdown(
            letterModelMd,
            "{% assign total = NoSuchThing %}{{ total }}");
        return Verify(result);
    }

    // forloop is scoped to a loop body in both validators, so outside one it stays an error.
    [Test]
    public Task Markdown_ForLoopIdentifierOutsideLoop_Reported()
    {
        var result = GeneratorDriver.RunMarkdown(letterModelMd, "{{ forloop.index }}");
        return Verify(result);
    }

    [Test]
    public Task Markdown_TemplateFileMissing()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("does-not-exist.md")]
            public partial class Empty;
            """;
        // No additional file added — pipeline must emit PARCH004.
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return Verify(result);
    }

    [Test]
    public async Task Markdown_InlineBlockTag_NoDiagnostic()
    {
        // Block tags inline with text are legal in markdown templates — Fluid parses the whole
        // file and there's no docx-style "block tag must sit alone in its paragraph" rule.
        var result = GeneratorDriver.RunMarkdown(letterModelMd, "Hello {% if Customer %}{{ Customer.Name }}{% endif %}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH005")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH001")).IsFalse();
    }

    [Test]
    public Task NestedClass_Valid()
    {
        // The decorated model sits inside a partial enclosing class. Generated source must wrap
        // the registration helper in `partial class Outer { partial class LetterModel { ... } }`.
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            public partial class Outer
            {
                [ParchmentModel("template.docx")]
                public partial class LetterModel
                {
                    public Customer Customer { get; set; } = new();
                }
            }
            """;
        var result = GeneratorDriver.Run(source, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public Task NestedClass_DoubleNested_Valid()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            public partial class Outer
            {
                public partial class Inner
                {
                    [ParchmentModel("template.docx")]
                    public partial class LetterModel
                    {
                        public Customer Customer { get; set; } = new();
                    }
                }
            }
            """;
        var result = GeneratorDriver.Run(source, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public async Task NestedClass_EnclosingNotPartial_Diagnostic()
    {
        // Outer is not partial. Should emit PARCH011 and skip generation entirely.
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            public class Outer
            {
                [ParchmentModel("template.docx")]
                public partial class LetterModel
                {
                    public Customer Customer { get; set; } = new();
                }
            }
            """;
        var result = GeneratorDriver.Run(source, "Hello {{ Customer.Name }}!");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Count(_ => _.Id == "PARCH011")).IsEqualTo(1);
        await Assert.That(diagnostics.Single(_ => _.Id == "PARCH011").GetMessage()).Contains("Outer");
        // Registration must NOT be emitted — the user would otherwise see a misleading CS0260
        // about "Missing partial modifier" with no link back to the SG's expectation.
        await Assert.That(result.GeneratedTrees.Length).IsEqualTo(0);
    }

    [Test]
    public Task NestedRecord_Valid()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            public partial record OuterRecord
            {
                [ParchmentModel("template.docx")]
                public partial class LetterModel
                {
                    public Customer Customer { get; set; } = new();
                }
            }
            """;
        var result = GeneratorDriver.Run(source, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public Task MixedInlineBlockTag()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                public string Description { get; set; } = "";
            }

            [ParchmentModel("template.docx")]
            public partial class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "prefix {% for line in Lines %}",
            "{{ line.Description }}",
            "{% endfor %}");
        return Verify(result);
    }

    [Test]
    public async Task MissingMember_NestedTarget_IncludesEnclosingChain()
    {
        // PARCH001 messages must disambiguate by enclosing chain. With seven `XxxGenerator.Info`
        // records in MinistersManager, a bare "not a member of 'Info'" was useless — the user
        // couldn't tell which Info the missing member was on. Now the message reads
        // "not a member of 'OuterGen.Info'".
        var source =
            """
            using Parchment;

            public static partial class OuterGen
            {
                [ParchmentModel("template.docx")]
                public partial class Info
                {
                    public string Name { get; set; } = "";
                }
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Missing }}");
        var diagnostic = result.Results.Single().Diagnostics.Single(_ => _.Id == "PARCH001");
        await Assert.That(diagnostic.GetMessage()).Contains("'OuterGen.Info'");
    }

    [Test]
    public async Task MultiTarget_SameSimpleName()
    {
        // Two `[ParchmentModel]` targets share the simple name `Info` but live inside different
        // enclosing classes. The SG must build distinct hint names from each target's enclosing
        // chain — otherwise `context.AddSource` throws "The hintName 'Info_ParchmentModel.g.cs'
        // must be unique within a generator" and the whole generator output is discarded.
        // Encountered when migrating MinistersManager, whose seven binding records are all
        // `XxxGenerator.Info`.
        var source =
            """
            using Parchment;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            public static partial class FirstGenerator
            {
                [ParchmentModel("first.docx")]
                public partial record Info
                {
                    public Customer Customer { get; set; } = new();
                }
            }

            public static partial class SecondGenerator
            {
                [ParchmentModel("second.docx")]
                public partial record Info
                {
                    public Customer Customer { get; set; } = new();
                }
            }
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            ("first.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!")),
            ("second.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!")));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();

        // No CS8785 "generator failed" diagnostic should land in the result's overall diagnostics
        // collection (that's where AddSource hint-name collisions surface).
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics).IsEmpty();
        // Both targets must produce a source file.
        await Assert.That(result.GeneratedTrees.Length).IsEqualTo(2);
        var hintNames = result.GeneratedTrees
            .Select(_ => Path.GetFileName(_.FilePath))
            .OrderBy(_ => _)
            .ToList();
        await Assert.That(hintNames[0]).IsEqualTo("FirstGenerator_Info_ParchmentModel.g.cs");
        await Assert.That(hintNames[1]).IsEqualTo("SecondGenerator_Info_ParchmentModel.g.cs");
    }

    [Test]
    public async Task RecordTarget_Valid()
    {
        // The decorated type is a record. The generator must emit `partial record` to match
        // the user's declaration — emitting `partial class` would fail with CS0261 "Partial
        // declarations have inconsistent type." Guards against regressing the kind-keyword
        // threading through TargetInfo.
        var source =
            """
            using Parchment;

            namespace Sample;

            public record Customer(string Name);

            [ParchmentModel("template.docx")]
            public partial record Letter
            {
                public required Customer Customer { get; init; }
            }
            """;
        var result = GeneratorDriver.Run(source, "Hello {{ Customer.Name }}!");
        var generated = result.GeneratedTrees.Single().ToString();
        await Assert.That(generated).Contains("partial record Letter");
        await Verify(result);
    }

    [Test]
    public Task AccessorEmission_AllFourMaps()
    {
        // Composite model exercising every emission shape: Fluid accessors for each reachable
        // type, an [ExcelsiorTable] collection, an [Html] string, a [Markdown] string, and an
        // IEnumerable<string> (auto bullet-list dispatch). The verified snapshot is the canonical
        // proof of what the SG produces.
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class HtmlAttribute : System.Attribute { }

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class MarkdownAttribute : System.Attribute { }

            public class Line
            {
                public string Description { get; set; } = "";
            }

            public class Customer
            {
                public string Name { get; set; } = "";
                public List<string> Tags { get; set; } = new();
            }

            [ParchmentModel("template.docx")]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();

                [ExcelsiorTable]
                public List<Line> Lines { get; set; } = new();

                [Html]
                public string Footer { get; set; } = "";

                [Markdown]
                public string Notes { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{{ Customer.Name }}",
            "{{ Customer.Tags }}",
            "{{ Lines }}",
            "{{ Footer }}",
            "{{ Notes }}");
        return Verify(result);
    }

    [Test]
    public async Task InheritedMember_Validates()
    {
        // ShapeBuilder.BuildEntry walks `current.BaseType` so members declared on a base class are
        // visible to validation. Without that walk, `{{ Title }}` would trip PARCH001 even though
        // the runtime resolves it via reflection.
        var source =
            """
            using Parchment;

            namespace Sample;

            public class DocumentBase
            {
                public string Title { get; set; } = "";
            }

            [ParchmentModel("template.docx")]
            public partial class Report : DocumentBase
            {
                public string Body { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Title }} — {{ Body }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics).IsEmpty();
    }

    // A render attribute on a static member is dropped: the maps that dispatch it walk instance
    // members only, so the value renders as plain text and the attribute does nothing. Silent at
    // runtime, so the generator says so.
    [Test]
    public async Task RenderAttributeOnStaticMemberWarns()
    {
        // The generator matches [Html] by type name, and it lives in Excelsior, which this
        // compilation does not reference — so it is declared locally, as the other tests do.
        var source =
            """
            using Parchment;

            namespace Sample;

            public sealed class HtmlAttribute : System.Attribute
            {
            }

            [ParchmentModel("template.docx")]
            public partial class Report
            {
                [Html]
                public static string Banner { get; set; } = "";

                public string Title { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Title }}");
        var diagnostics = result.Results.Single().Diagnostics;

        var warning = diagnostics.SingleOrDefault(_ => _.Id == "PARCH019");
        await Assert.That(warning).IsNotNull();
        await Assert.That(warning!.GetMessage()).Contains("Banner");
        await Assert.That(warning.GetMessage()).Contains("Html");
    }

    [Test]
    public async Task RenderAttributeOnInstanceMemberDoesNotWarn()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public sealed class HtmlAttribute : System.Attribute
            {
            }

            [ParchmentModel("template.docx")]
            public partial class Report
            {
                [Html]
                public string Banner { get; set; } = "";

                public string Title { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Title }} {{ Banner }}");
        var diagnostics = result.Results.Single().Diagnostics;

        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH019")).IsFalse();
    }

    // A plain static member carries no attribute to drop, so it is not worth a warning.
    [Test]
    public async Task PlainStaticMemberDoesNotWarn()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Report
            {
                public static string Banner { get; set; } = "";

                public string Title { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Title }} {{ Banner }}");
        var diagnostics = result.Results.Single().Diagnostics;

        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH019")).IsFalse();
    }
}
