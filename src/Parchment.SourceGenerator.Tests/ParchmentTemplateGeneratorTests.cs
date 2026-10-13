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

        [ParchmentModel]
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

        [ParchmentModel]
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

        [ParchmentModel]
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

            [ParchmentModel]
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

            [ParchmentModel]
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

    const string rangeModelDocx =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Doc
        {
            public int Count { get; set; }

            public List<Line> Lines { get; set; } = new();
        }
        """;

    // A range has no member path, so the loop variable is untyped: it exists, and members reached
    // through it are not checked. The generator used to skip the tag without opening a loop, which
    // left `{{ i }}` an unknown identifier and paired the `{% endfor %}` with an enclosing loop.
    [Test]
    public async Task ForLoop_Range_NoDiagnostic()
    {
        foreach (var range in new[] { "(1..3)", "(1..Count)" })
        {
            var result = GeneratorDriver.Run(
                rangeModelDocx,
                $"{{% for i in {range} %}}",
                "{{ i }} {{ i.Anything }}",
                "{% if i.Anything %}",
                "x",
                "{% endif %}",
                "{% for inner in i.Anything %}",
                "{{ inner.Anything }}",
                "{% endfor %}",
                "{% endfor %}");

            await Assert.That(result.Results.Single().Diagnostics).IsEmpty();
        }
    }

    // A member used as a bound is a reference like any other.
    [Test]
    public async Task ForLoop_Range_MissingMemberBound_ReportsParch001()
    {
        var result = GeneratorDriver.Run(
            rangeModelDocx,
            "{% for i in (1..NoSuchMember) %}",
            "{{ i }}",
            "{% endfor %}");

        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Length).IsEqualTo(1);
        await Assert.That(diagnostics[0].Id).IsEqualTo("PARCH001");
        await Assert.That(diagnostics[0].GetMessage()).Contains("NoSuchMember");
    }

    [Test]
    public async Task ForLoop_Range_VariableDoesNotLeakAfterEndfor()
    {
        var result = GeneratorDriver.Run(
            rangeModelDocx,
            "{% for i in (1..3) %}",
            "{{ i }}",
            "{% endfor %}",
            "{{ i }}");

        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Length).IsEqualTo(1);
        await Assert.That(diagnostics[0].Id).IsEqualTo("PARCH001");
    }

    // The range's `{% endfor %}` must close the range loop, not the enclosing one: `line` stays
    // bound until its own `{% endfor %}`.
    [Test]
    public async Task ForLoop_RangeNestedInTypedLoop_EndforPairsWithRange()
    {
        var result = GeneratorDriver.Run(
            rangeModelDocx,
            "{% for line in Lines %}",
            "{% for i in (1..3) %}",
            "{{ i }} {{ line.Description }}",
            "{% endfor %}",
            "{{ line.Description }}",
            "{% endfor %}");

        await Assert.That(result.Results.Single().Diagnostics).IsEmpty();

        var typo = GeneratorDriver.Run(
            rangeModelDocx,
            "{% for i in (1..3) %}",
            "{% for line in Lines %}",
            "{{ line.NoSuchMember }}",
            "{% endfor %}",
            "{% endfor %}");

        var diagnostics = typo.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Length).IsEqualTo(1);
        await Assert.That(diagnostics[0].GetMessage()).Contains("line.NoSuchMember");
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

            [ParchmentModel]
            public partial class Empty;
            """;
        var result = GeneratorDriver.Run(source, "{% foobar %}");
        return Verify(result);
    }

    // PARCH003 used to fire for every dashed tag, on a template the runtime registers and renders.
    [Test]
    public async Task WhitespaceControlOnBlockTags_ReportsNothing()
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

            [ParchmentModel]
            public partial class Invoice
            {
                public List<Line> Lines { get; set; } = new();
                public bool Flag { get; set; }
                public bool Other { get; set; }
                public string Name { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{%- for line in Lines -%}",
            "{{- line.Description -}}",
            "{%-endfor-%}",
            "{%- for line in Lines %}",
            "{{ line.Description -}}",
            "{% endfor -%}",
            "{%- if Flag -%}",
            "{{- Name }}",
            "{%- elsif Other -%}",
            "b",
            "{%- else -%}",
            "c",
            "{%- endif -%}");
        await Assert.That(result.Results.Single().Diagnostics).IsEmpty();
    }

    // A missing member behind a dashed tag is still found, so the hyphens hide nothing.
    [Test]
    public async Task WhitespaceControlOnBlockTags_StillValidatesTheExpression()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class Empty;
            """;
        var result = GeneratorDriver.Run(
            source,
            "{%- if Missing -%}",
            "a",
            "{%- endif -%}",
            "{%- foobar -%}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEquivalentTo(["PARCH001", "PARCH003"]);
    }

    [Test]
    public Task TemplateFileMissing()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class Empty;
            """;
        // The only AdditionalFile is named after some other type, so the convention finds nothing
        // for Empty and must emit PARCH004.
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Other.parchment.docx", GeneratorDriver.BuildDocxBytes("ignored")));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return Verify(result);
    }

    [Test]
    public async Task BothDocxAndMarkdownMatch_Ambiguous()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class Empty;
            """;
        // Two AdditionalFiles carry the type's name — a docx and an md. The generator reports
        // rather than picking one.
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Empty.parchment.docx", GeneratorDriver.BuildDocxBytes("hello")),
            new TemplateFile("Empty.parchment.md", "hello"u8.ToArray()));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Single().Id).IsEqualTo("PARCH020");
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

            [ParchmentModel]
            public partial class Letter
            {
                public Customer Customer { get; set; } = new();
            }

            [ParchmentModel]
            public partial class Invoice
            {
                public decimal Total { get; set; }
            }
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Letter.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!")),
            new TemplateFile("Invoice.parchment.docx", GeneratorDriver.BuildDocxBytes("Total: {{ Total }}")));

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

            [ParchmentModel]
            public partial class Empty;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Empty.parchment.docx", "not a zip file"u8.ToArray()));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var diagnostics = result.Results.Single().Diagnostics;

        await Assert.That(diagnostics.Length).IsEqualTo(1);
        await Assert.That(diagnostics[0].Id).IsEqualTo("PARCH006");
    }

    // A word processor keeps the document it is editing open with write access, whichever way it
    // shares it. The generator only reads, so that is not a PARCH006: the template is validated
    // (the PARCH001 below) and the bytes on disk are the ones embedded.
    [Test]
    [Arguments(FileShare.Read)]
    [Arguments(FileShare.ReadWrite)]
    public async Task DocxOpenForWritingElsewhere_IsStillRead(FileShare share)
    {
        var bytes = GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!", "{{ Missing }}");
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            letterModelDocx,
            new TemplateFile("Letter.parchment.docx", bytes));

        GeneratorDriverRunResult result;
        using (new FileStream(setup.DocxPath, FileMode.Open, FileAccess.ReadWrite, share))
        {
            result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        }

        await Assert.That(result.Diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH001"]);
        await Assert.That(GeneratedText(result)).Contains(Convert.ToBase64String(bytes));
    }

    // Saving replaces the file, which takes a handle with delete access. DeleteOnClose is the
    // way to hold one from here.
    [Test]
    public async Task DocxOpenForDeletionElsewhere_IsStillRead()
    {
        var bytes = GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!", "{{ Missing }}");
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            letterModelDocx,
            new TemplateFile("Letter.parchment.docx", bytes));

        GeneratorDriverRunResult result;
        using (new FileStream(
                   setup.DocxPath,
                   FileMode.Open,
                   FileAccess.ReadWrite,
                   FileShare.ReadWrite | FileShare.Delete,
                   bufferSize: 1,
                   FileOptions.DeleteOnClose))
        {
            result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        }

        await Assert.That(result.Diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH001"]);
        await Assert.That(GeneratedText(result)).Contains(Convert.ToBase64String(bytes));
    }

    // The style document of a markdown template is read the same way as a docx template.
    [Test]
    [Arguments(FileShare.Read)]
    [Arguments(FileShare.ReadWrite)]
    public async Task StyleDocOpenForWritingElsewhere_IsStillRead(FileShare share)
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var setup = GeneratorDriver.CreateDriverWithFiles(
            letterModelMd,
            new TemplateFile("Letter.parchment.md", "Hello {{ Customer.Name }}!"u8.ToArray()),
            new TemplateFile("Letter.parchment.dotx", styleBytes));

        GeneratorDriverRunResult result;
        using (new FileStream(setup.DocxPaths[1], FileMode.Open, FileAccess.ReadWrite, share))
        {
            result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        }

        await Assert.That(result.Diagnostics).IsEmpty();
        await Assert.That(GeneratedText(result)).Contains(Convert.ToBase64String(styleBytes));
    }

    [Test]
    public async Task StyleDocOpenForDeletionElsewhere_IsStillRead()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var setup = GeneratorDriver.CreateDriverWithFiles(
            letterModelMd,
            new TemplateFile("Letter.parchment.md", "Hello {{ Customer.Name }}!"u8.ToArray()),
            new TemplateFile("Letter.parchment.dotx", styleBytes));

        GeneratorDriverRunResult result;
        using (new FileStream(
                   setup.DocxPaths[1],
                   FileMode.Open,
                   FileAccess.ReadWrite,
                   FileShare.ReadWrite | FileShare.Delete,
                   bufferSize: 1,
                   FileOptions.DeleteOnClose))
        {
            result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        }

        await Assert.That(result.Diagnostics).IsEmpty();
        await Assert.That(GeneratedText(result)).Contains(Convert.ToBase64String(styleBytes));
    }

    static string GeneratedText(GeneratorDriverRunResult result) =>
        string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

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

            [ParchmentModel]
            public partial class Empty;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Empty.parchment.docx", GeneratorDriver.BuildDocxBytesWithoutPrivacyFlag("hello")));

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

            [ParchmentModel]
            public partial class Empty;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Empty.parchment.docx", GeneratorDriver.BuildDocxBytes("hello")));

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

        [ParchmentModel]
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

    // Lockstep with the runtime's ExcelsiorTokenValidator: both ask TokenScan.IsSolo, so padding
    // around the token is fine and a second token in the paragraph is not.
    [Test]
    [Arguments("{{ Lines }} ")]
    [Arguments(" {{ Lines }}")]
    public async Task ExcelsiorToken_PaddedWithWhitespace_NoDiagnostics(string paragraph)
    {
        var result = GeneratorDriver.Run(excelsiorModel, paragraph);
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsFalse();
    }

    [Test]
    [Arguments("{{ Lines }}{{ Lines }}")]
    [Arguments("{{ Lines }} {{ Lines }}")]
    public async Task ExcelsiorToken_SharingParagraphWithAnotherToken_Diagnostic(string paragraph)
    {
        var result = GeneratorDriver.Run(excelsiorModel, paragraph);
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

    // The markdown flow rewrites the token to a marker before the source is parsed, so the same two
    // rules apply there — read off the line rather than off the paragraph. Without these, the
    // mistake surfaces when the store materializes the template on first render, not at build.
    [Test]
    public async Task Markdown_ExcelsiorToken_SharingItsLine_Diagnostic()
    {
        var result = GeneratorDriver.RunMarkdown(excelsiorModel, "Lines: {{ Lines }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsTrue();
    }

    [Test]
    public async Task Markdown_ExcelsiorToken_WithFilter_Diagnostic()
    {
        var result = GeneratorDriver.RunMarkdown(excelsiorModel, "{{ Lines | reverse }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH008")).IsTrue();
    }

    [Test]
    public async Task Markdown_ExcelsiorToken_Clean_NoDiagnostics()
    {
        var result = GeneratorDriver.RunMarkdown(excelsiorModel, "# Invoice\n\n{{ Lines }}\n");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH008")).IsFalse();
    }

    // A loop variable's members are not reachable from the root model, so the table lookup misses
    // and the token falls through to Fluid — the same as the docx flow. Not a diagnostic.
    [Test]
    public async Task Markdown_ExcelsiorToken_LoopScoped_NoDiagnostics()
    {
        var result = GeneratorDriver.RunMarkdown(
            excelsiorModel,
            "{% for line in Lines %}\n{{ line.Description }}\n{% endfor %}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH008")).IsFalse();
    }

    // `order` is both a loop variable and, but for case, a root member whose own members carry
    // the attributes. The loop item's members of the same names carry none.
    const string shadowedModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute { }

        public class Line
        {
            public string Description { get; set; } = "";
        }

        public class Order
        {
            [ExcelsiorTable]
            public List<Line> Lines { get; set; } = new();

            [Html]
            public string Notes { get; set; } = "";
        }

        public class Summary
        {
            public int Lines { get; set; }

            public string Notes { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Invoice
        {
            public Order Order { get; set; } = new();

            public List<Summary> Orders { get; set; } = new();
        }
        """;

    static readonly string[] attributeRuleIds = ["PARCH007", "PARCH008", "PARCH010", "PARCH024"];

    // Inside the loop the tokens read the item, so none of the root members' rules apply: not
    // alone-in-its-paragraph, not plain-member-access, and no marker is being lost. The runtime
    // agrees — see LoopTests.LoopVariableNamedAfterRoot*.
    [Test]
    public async Task LoopVariableNamedAfterRootMember_NoAttributeDiagnostics()
    {
        var result = GeneratorDriver.Run(
            shadowedModel,
            "{% for order in Orders %}",
            "Next: {{ order.Lines | plus: 1 }}",
            "{{ order.Notes | upcase }}",
            "{% endfor %}");

        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Intersect(attributeRuleIds)).IsEmpty();
    }

    [Test]
    public async Task LoopVariableNamedAfterRootMember_NoAttributeDiagnostics_Markdown()
    {
        var result = GeneratorDriver.RunMarkdown(
            shadowedModel,
            "{% for order in Orders %}\nNext: {{ order.Lines | plus: 1 }}\n\n{{ order.Notes | upcase }}\n{% endfor %}");

        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Intersect(attributeRuleIds)).IsEmpty();
    }

    // Outside the loop the same spelling is the root member again.
    [Test]
    public async Task RootMemberAfterLoopOfTheSameName_StillReported_Markdown()
    {
        var result = GeneratorDriver.RunMarkdown(
            shadowedModel,
            "{% for order in Orders %}\n{{ order.Notes }}\n{% endfor %}\nLines: {{ Order.Lines }}");

        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsTrue();
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

        [ParchmentModel]
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

    // The attribute types are deliberately not declared. A source generator sees only the input
    // compilation, so an attribute another generator emits — StringSyntaxAttributeAnalyzer's [Html]
    // shortcut, for one — reaches this one as an error symbol named exactly as written: "Html", not
    // "HtmlAttribute". Matching only the declared spelling dropped the marker without a word, and the
    // model rendered its html as visible tags.
    const string generatedMarkerModel =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Doc
        {
            [Html]
            public string Body { get; set; } = "";

            [Markdown]
            public string Notes { get; set; } = "";
        }
        """;

    // The same model with the markers removed, so the test below pins the marker rather than
    // something every string member would get.
    const string unmarkedModel =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Doc
        {
            public string Body { get; set; } = "";

            public string Notes { get; set; } = "";
        }
        """;

    static string Emitted(GeneratorDriverRunResult result) =>
        string.Concat(result.Results.SelectMany(_ => _.GeneratedSources).Select(_ => _.SourceText.ToString()));

    [Test]
    public async Task FormatMarker_FromAnotherGenerator_IsStillHonoured()
    {
        var emitted = Emitted(GeneratorDriver.Run(generatedMarkerModel, "{{ Body }}", "{{ Notes }}"));

        await Assert.That(emitted).Contains("new(\"Body\", global::Parchment.Generated.FormatMapKind.Html");
        await Assert.That(emitted).Contains("new(\"Notes\", global::Parchment.Generated.FormatMapKind.Markdown");
    }

    [Test]
    public async Task UnmarkedMembers_EmitNoFormatEntries()
    {
        var emitted = Emitted(GeneratorDriver.Run(unmarkedModel, "{{ Body }}", "{{ Notes }}"));

        await Assert.That(emitted).DoesNotContain("FormatMapKind.Html");
        await Assert.That(emitted).DoesNotContain("FormatMapKind.Markdown");
    }

    // A format marker on a loop element. The runtime resolves one by walking the model from its
    // root, so it cannot address the item an iteration is on — the marker is silently not applied
    // and the markup shows in the document.
    const string loopedMarkerModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute { }

        public class Line
        {
            [Html]
            public string Body { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Doc
        {
            public List<Line> Lines { get; set; } = new();

            [Html]
            public string Intro { get; set; } = "";
        }
        """;

    [Test]
    public async Task FormatMarker_ThroughLoopVariable_ReportsParch024()
    {
        var result = GeneratorDriver.Run(
            loopedMarkerModel,
            "{% for line in Lines %}",
            "{{ line.Body }}",
            "{% endfor %}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH024")).IsTrue();
    }

    // Addressable from the root, so the marker is honoured and nothing is reported.
    [Test]
    public async Task FormatMarker_FromRoot_NoParch024()
    {
        var result = GeneratorDriver.Run(loopedMarkerModel, "{{ Intro }}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH024")).IsFalse();
    }

    // An unmarked member reached the same way is ordinary text — the loop is not what is wrong.
    [Test]
    public async Task UnmarkedMember_ThroughLoopVariable_NoParch024()
    {
        var source = loopedMarkerModel.Replace("    [Html]", "    ");
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "{{ line.Body }}",
            "{% endfor %}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH024")).IsFalse();
    }

    // The markdown flow reaches the same conclusion by a different route: the rewrite that applies
    // a marker matches a token path against the map built from the model root, and a loop variable
    // is not addressable from there.
    [Test]
    public async Task FormatMarker_ThroughLoopVariable_ReportsParch024_Markdown()
    {
        var result = GeneratorDriver.RunMarkdown(
            loopedMarkerModel,
            "{% for line in Lines %}\n{{ line.Body }}\n{% endfor %}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH024")).IsTrue();
    }

    [Test]
    public async Task FormatMarker_FromRoot_NoParch024_Markdown()
    {
        var result = GeneratorDriver.RunMarkdown(loopedMarkerModel, "{{ Intro }}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH024")).IsFalse();
    }

    // A member that declares its rendering twice: once by type, once by marker. Both are applied,
    // and the second converts the placeholder the first produced — so the placeholder is what lands
    // in the document.
    const string markerOnTokenModel =
        """
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute { }

        [ParchmentModel]
        public partial class Doc
        {
            [Html]
            public HtmlToken Body => new("");
        }
        """;

    [Test]
    public async Task FormatMarkerOnTokenValue_ReportsParch025()
    {
        var result = GeneratorDriver.Run(markerOnTokenModel, "{{ Body }}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH025")).IsTrue();
    }

    // The conflicted member is left out of the format map, so nothing applies the marker on top of
    // the type while the build is failing over it.
    [Test]
    public async Task FormatMarkerOnTokenValue_EmitsNoFormatEntry()
    {
        var emitted = Emitted(GeneratorDriver.Run(markerOnTokenModel, "{{ Body }}"));

        await Assert.That(emitted).DoesNotContain("FormatMapKind.Html");
    }

    // The type on its own is the supported way to say it, and draws nothing.
    [Test]
    public async Task TokenValueWithoutMarker_NoParch025()
    {
        var source = markerOnTokenModel.Replace("    [Html]", "    ");
        var result = GeneratorDriver.Run(source, "{{ Body }}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH025")).IsFalse();
    }

    // So is the marker on its own, on a plain string.
    [Test]
    public async Task MarkerOnPlainString_NoParch025()
    {
        var source = markerOnTokenModel.Replace("public HtmlToken Body => new(\"\");", "public string Body => \"\";");
        var result = GeneratorDriver.Run(source, "{{ Body }}");

        await Assert.That(result.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH025")).IsFalse();
    }

    // A marker on a member that is not a string. The conversion reads a string, so the value used
    // to render as nothing at all.
    const string markerOnNonStringModel =
        """
        using System.Diagnostics.CodeAnalysis;
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute { }

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class MarkdownAttribute : System.Attribute { }

        public class Rich
        {
            public string Text { get; set; } = "";
        }

        public class Section
        {
            [Markdown]
            public System.DateTime When { get; set; }
        }

        [ParchmentModel]
        public partial class Doc
        {
            [Html]
            public int Count { get; set; }

            [Html]
            public Rich Rich { get; set; } = new();

            [Markdown]
            public object Body { get; set; } = 42;

            [Html]
            public decimal? Price { get; set; }

            [StringSyntax("html")]
            public System.Uri? Link { get; set; }

            public Section Section { get; set; } = new();

            [Html]
            public string? Fine { get; set; }
        }
        """;

    static readonly string[] nonStringMembers =
    [
        "Doc.Count",
        "Doc.Rich",
        "Doc.Body",
        "Doc.Price",
        "Doc.Link",
        "Section.When"
    ];

    static async Task AssertParch029(GeneratorDriverRunResult result)
    {
        var messages = result.Results.Single().Diagnostics
            .Where(_ => _.Id == "PARCH029")
            .Select(_ => _.GetMessage())
            .ToList();

        await Assert.That(messages.Count).IsEqualTo(nonStringMembers.Length);
        foreach (var member in nonStringMembers)
        {
            await Assert.That(messages.Any(_ => _.Contains($"'Sample.{member}'"))).IsTrue().Because(member);
        }
    }

    [Test]
    public async Task FormatMarkerOnNonString_ReportsParch029() =>
        await AssertParch029(GeneratorDriver.Run(markerOnNonStringModel, "{{ Count }}"));

    // A markdown template reads the marker too: it rewrites the token to carry the filter.
    [Test]
    public async Task FormatMarkerOnNonString_ReportsParch029_Markdown() =>
        await AssertParch029(GeneratorDriver.RunMarkdown(markerOnNonStringModel, "{{ Count }}"));

    [Test]
    public async Task FormatMarkerOnNonString_ReportsParch029_Bindable()
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(markerOnNonStringModel.Replace("[ParchmentModel]", "[ParchmentBindable]"));

        await AssertParch029(setup.Driver.RunGenerators(setup.Compilation).GetRunResult());
    }

    // Only the string member reaches the format map, so nothing reads the others as a string
    // while the build is failing over them.
    [Test]
    public async Task FormatMarkerOnNonString_EmitsNoFormatEntry()
    {
        var emitted = Emitted(GeneratorDriver.Run(markerOnNonStringModel, "{{ Count }}"));

        await Assert.That(emitted).Contains("new(\"Fine\", global::Parchment.Generated.FormatMapKind.Html");
        await Assert.That(emitted).DoesNotContain("FormatMapKind.Markdown");
        foreach (var name in new[] {"Count", "Rich", "Price", "Link"})
        {
            await Assert.That(emitted).DoesNotContain($"new(\"{name}\", global::Parchment.Generated.FormatMapKind");
        }
    }

    // A nullable string is still a string, and a token type is PARCH025's to report.
    [Test]
    public async Task MarkerOnStringOrToken_NoParch029()
    {
        var nullable = markerOnTokenModel.Replace("public HtmlToken Body => new(\"\");", "public string? Body => null;");
        var onString = GeneratorDriver.Run(nullable, "{{ Body }}");
        var onToken = GeneratorDriver.Run(markerOnTokenModel, "{{ Body }}");

        await Assert.That(onString.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH029")).IsFalse();
        await Assert.That(onToken.Results.Single().Diagnostics.Any(_ => _.Id == "PARCH029")).IsFalse();
    }

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

            [ParchmentModel]
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

            [ParchmentModel]
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

    // Fluid rejects a filter on a for source, so it never reaches validation. Pinned because the
    // shape looks supportable and the validators were once thought to differ over it.
    [Test]
    public Task Markdown_FilteredForLoopSource_IsAParseError()
    {
        var result = GeneratorDriver.RunMarkdown(
            invoiceModelMd,
            """
            {% for line in Lines | reverse %}
            - {{ line.Description }}
            {% endfor %}
            """);
        return Verify(result);
    }

    // A range has no member path at all, so the loop variable is untyped rather than bound to the
    // root model. Binding to the root reported a PARCH001 on every member reached through it.
    [Test]
    public Task Markdown_RangeForLoopSource_NoDiagnostic()
    {
        var result = GeneratorDriver.RunMarkdown(
            letterModelMd,
            """
            {% for i in (1..5) %}
            {{ i.size }}
            {% endfor %}
            """);
        return Verify(result);
    }

    [Test]
    public Task Markdown_TemplateFileMissing()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
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
                [ParchmentModel]
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
                    [ParchmentModel]
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
                [ParchmentModel]
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
                [ParchmentModel]
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

            [ParchmentModel]
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
                [ParchmentModel]
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
                [ParchmentModel]
                public partial record Info
                {
                    public Customer Customer { get; set; } = new();
                }
            }

            public static partial class SecondGenerator
            {
                [ParchmentModel]
                public partial record Info
                {
                    public Customer Customer { get; set; } = new();
                }
            }
            """;

        // Both targets share the simple name Info, so under the convention they bind the same
        // template file — each gets its own embedded copy.
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Info.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Customer.Name }}!")));

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

            [ParchmentModel]
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

            [ParchmentModel]
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
    public Task ExcelsiorConfigure_EmitsDirectCall()
    {
        // The generated entry closes over the element type and calls the named method directly, so
        // a wrong signature is an ordinary compile error rather than a render-time failure. The
        // method takes object here because this compilation does not reference Excelsior - the
        // typed cast in the emission is what the snapshot pins.
        var source =
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
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                internal static void ConfigureLines(object builder)
                {
                }
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Lines }}");
        return Verify(result);
    }

    // The generated code calls the method directly, so an unresolvable name would surface as a
    // compile error inside a .g.cs file - correct, but pointing at code nobody wrote. Reported on
    // the member that named it instead, with the emission suppressed so this is the only error.
    [Test]
    public async Task ExcelsiorConfigureMissing_Reports()
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

            [ParchmentModel]
            public partial class Invoice
            {
                [ExcelsiorTable(Configure = "Nope")]
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Lines }}");
        var diagnostics = result.Results.Single().Diagnostics;

        var error = diagnostics.SingleOrDefault(_ => _.Id == "PARCH026");
        await Assert.That(error).IsNotNull();
        await Assert.That(error!.GetMessage()).Contains("Nope");
        var generated = result.Results.Single().GeneratedSources
            .Single(_ => _.HintName.Contains("Invoice"));
        await Assert.That(generated.SourceText.ToString()).DoesNotContain("Nope");
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

            [ParchmentModel]
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

            [ParchmentModel]
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

            [ParchmentModel]
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

            [ParchmentModel]
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

    const string pagedModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Paged
        {
            public List<Line> Lines { get; set; } = new();
            public int Max { get; set; }
        }
        """;

    static List<string> PagedDiagnostics(string openTag) =>
        GeneratorDriver.Run(pagedModel, openTag, "{{ line.Description }}", "{% endfor %}")
            .Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .Where(_ => _.StartsWith("PARCH", StringComparison.Ordinal))
            .ToList();

    // The runtime evaluates limit and offset as expressions, so a member named there is a model
    // reference like any other and is validated at registration.
    [Test]
    [Arguments("{% for line in Lines limit: Max %}")]
    [Arguments("{% for line in Lines offset: Max %}")]
    [Arguments("{% for line in Lines limit:2 offset:1 reversed %}")]
    [Arguments("{% for line in Lines limit: Lines.size %}")]
    public async Task Docx_LoopModifier_Valid_NoDiagnostic(string openTag) =>
        await Assert.That(PagedDiagnostics(openTag)).IsEmpty();

    [Test]
    [Arguments("{% for line in Lines limit: Missing %}")]
    [Arguments("{% for line in Lines offset: Missing %}")]
    // The CLR member is not what liquid binds; `size` is.
    [Arguments("{% for line in Lines limit: Lines.Count %}")]
    // Evaluated before the loop variable is bound.
    [Arguments("{% for line in Lines limit: line %}")]
    // The docx flow has no liquid loop scope to resume from; the runtime rejects it too.
    [Arguments("{% for line in Lines offset: continue %}")]
    public async Task Docx_LoopModifier_UnknownMember_ReportsParch001(string openTag) =>
        await Assert.That(PagedDiagnostics(openTag)).IsEquivalentTo(["PARCH001"]);

    static List<string> RangeModifierDiagnostics(string openTag) =>
        GeneratorDriver.Run(pagedModel, openTag, "{{ i }}", "{% endfor %}")
            .Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .ToList();

    // A range loop takes the modifiers like any other, and their references are validated on its
    // path too.
    [Test]
    [Arguments("{% for i in (1..5) limit:2 %}")]
    [Arguments("{% for i in (1..5) offset:1 %}")]
    [Arguments("{% for i in (1..5) reversed %}")]
    [Arguments("{% for i in (1..5) limit:2 offset:1 reversed %}")]
    [Arguments("{% for i in (1..Max) limit: Max offset: Lines.size %}")]
    public async Task Docx_RangeLoopModifier_Valid_NoDiagnostic(string openTag) =>
        await Assert.That(RangeModifierDiagnostics(openTag)).IsEmpty();

    [Test]
    [Arguments("{% for i in (1..5) limit: Missing %}")]
    [Arguments("{% for i in (1..5) offset: Missing %}")]
    [Arguments("{% for i in (1..5) limit: Lines.Count %}")]
    // Evaluated before the loop variable is bound, untyped or not.
    [Arguments("{% for i in (1..5) limit: i %}")]
    [Arguments("{% for i in (1..5) offset: continue %}")]
    public async Task Docx_RangeLoopModifier_UnknownMember_ReportsParch001(string openTag) =>
        await Assert.That(RangeModifierDiagnostics(openTag)).IsEquivalentTo(["PARCH001"]);
}
