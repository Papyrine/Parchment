public class ParchmentTemplateGeneratorTests
{
    const string letterModel =
        """
        using Parchment;

        namespace Sample;

        public class Customer
        {
            public string Name { get; set; } = "";
        }

        public class Letter
        {
            public Customer Customer { get; set; } = new();
        }
        """;

    [Test]
    public Task Substitution_Valid()
    {
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Letter))]
                     public partial class CustomerLetter;
                     """;
        var result = GeneratorDriver.Run(source, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public Task Substitution_MissingMember()
    {
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Letter))]
                     public partial class CustomerLetter;
                     """;
        var result = GeneratorDriver.Run(source, "{{ Customer.Missing }}");
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

            public class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentTemplate("template.docx", typeof(Invoice))]
            public partial class InvoiceDoc;
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "{{ line.Description }}",
            "{% endfor %}");
        return Verify(result);
    }

    [Test]
    public Task ForLoop_SourceNotEnumerable()
    {
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Letter))]
                     public partial class BadLoop;
                     """;
        var result = GeneratorDriver.Run(
            source,
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

            public class Empty;

            [ParchmentTemplate("template.docx", typeof(Empty))]
            public partial class WeirdTag;
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

            public class Empty;

            [ParchmentTemplate("does-not-exist.docx", typeof(Empty))]
            public partial class MissingFile;
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
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Letter))]
                     public partial class StringLoop;
                     """;
        var result = GeneratorDriver.Run(
            source,
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

            public class Letter
            {
                public Customer Customer { get; set; } = new();
            }

            public class Invoice
            {
                public decimal Total { get; set; }
            }

            [ParchmentTemplate("letter.docx", typeof(Letter))]
            public partial class LetterDoc;

            [ParchmentTemplate("invoice.docx", typeof(Invoice))]
            public partial class InvoiceDoc;
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

            public class Empty;

            [ParchmentTemplate("template.docx", typeof(Empty))]
            public partial class Corrupt;
            """;

        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            ("template.docx", "not a zip file"u8.ToArray()));

        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var diagnostics = result.Results.Single().Diagnostics;

        await Assert.That(diagnostics.Length).IsEqualTo(1);
        await Assert.That(diagnostics[0].Id).IsEqualTo("PARCH006");
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

        public class Invoice
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
        var source = excelsiorModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Invoice))]
                     public partial class MixedExcelsior;
                     """;
        var result = GeneratorDriver.Run(source, "Prefix {{ Lines }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH007")).IsTrue();
    }

    [Test]
    public async Task ExcelsiorToken_WithFilter_Diagnostic()
    {
        // Filter chain on an [ExcelsiorTable] substitution must trip PARCH008.
        var source = excelsiorModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Invoice))]
                     public partial class FilteredExcelsior;
                     """;
        var result = GeneratorDriver.Run(source, "{{ Lines | reverse }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH008")).IsTrue();
    }

    [Test]
    public async Task ExcelsiorToken_Clean_NoDiagnostics()
    {
        // Baseline: plain {{ Lines }} in its own paragraph should NOT trip PARCH007/PARCH008.
        var source = excelsiorModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Invoice))]
                     public partial class CleanExcelsior;
                     """;
        var result = GeneratorDriver.Run(source, "{{ Lines }}");
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

        public class Doc
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
        var source = formatModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Doc))]
                     public partial class MixedFormat;
                     """;
        var result = GeneratorDriver.Run(source, "Prefix {{ Body }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH009")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsFalse();
    }

    [Test]
    public async Task FormatToken_WithFilter_Diagnostic()
    {
        var source = formatModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Doc))]
                     public partial class FilteredFormat;
                     """;
        var result = GeneratorDriver.Run(source, "{{ Body | upcase }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsTrue();
    }

    [Test]
    public async Task FormatToken_MarkdownClean_NoDiagnostics()
    {
        var source = formatModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Doc))]
                     public partial class CleanFormat;
                     """;
        var result = GeneratorDriver.Run(source, "{{ Notes }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsFalse();
    }

    [Test]
    public async Task FormatToken_StringSyntaxHtml_NoDiagnostics()
    {
        var source = formatModel +
                     """

                     [ParchmentTemplate("template.docx", typeof(Doc))]
                     public partial class SyntaxFormat;
                     """;
        var result = GeneratorDriver.Run(source, "{{ Summary }}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH009")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH010")).IsFalse();
    }

    [Test]
    public Task Markdown_Substitution_Valid()
    {
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.md", typeof(Letter))]
                     public partial class CustomerLetterMd;
                     """;
        var result = GeneratorDriver.RunMarkdown(source, "Hello {{ Customer.Name }}!");
        return Verify(result);
    }

    [Test]
    public Task Markdown_Substitution_MissingMember()
    {
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.md", typeof(Letter))]
                     public partial class CustomerLetterMd;
                     """;
        var result = GeneratorDriver.RunMarkdown(source, "{{ Customer.Missing }}");
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

            public class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentTemplate("template.md", typeof(Invoice))]
            public partial class InvoiceMd;
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

            public class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentTemplate("template.md", typeof(Invoice))]
            public partial class InvoiceMd;
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
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.md", typeof(Letter))]
                     public partial class BadLoopMd;
                     """;
        var result = GeneratorDriver.RunMarkdown(
            source,
            """
            {% for item in Customer %}
            x
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

            public class Empty;

            [ParchmentTemplate("does-not-exist.md", typeof(Empty))]
            public partial class MissingMd;
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
        var source = letterModel +
                     """

                     [ParchmentTemplate("template.md", typeof(Letter))]
                     public partial class InlineBlockMd;
                     """;
        var result = GeneratorDriver.RunMarkdown(source, "Hello {% if Customer %}{{ Customer.Name }}{% endif %}");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH005")).IsFalse();
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH001")).IsFalse();
    }

    [Test]
    public Task NestedClass_Valid()
    {
        // The decorated class sits inside a partial enclosing class. Generated source must wrap
        // the registration helper in `partial class Outer { partial class Letter { ... } }`.
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            public class LetterModel
            {
                public Customer Customer { get; set; } = new();
            }

            public partial class Outer
            {
                [ParchmentTemplate("template.docx", typeof(LetterModel))]
                public partial class CustomerLetter;
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

            public class LetterModel
            {
                public Customer Customer { get; set; } = new();
            }

            public partial class Outer
            {
                public partial class Inner
                {
                    [ParchmentTemplate("template.docx", typeof(LetterModel))]
                    public partial class CustomerLetter;
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

            public class LetterModel
            {
                public Customer Customer { get; set; } = new();
            }

            public class Outer
            {
                [ParchmentTemplate("template.docx", typeof(LetterModel))]
                public partial class CustomerLetter;
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

            public class LetterModel
            {
                public Customer Customer { get; set; } = new();
            }

            public partial record OuterRecord
            {
                [ParchmentTemplate("template.docx", typeof(LetterModel))]
                public partial class CustomerLetter;
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

            public class Invoice
            {
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentTemplate("template.docx", typeof(Invoice))]
            public partial class Mixed;
            """;
        var result = GeneratorDriver.Run(
            source,
            "prefix {% for line in Lines %}",
            "{{ line.Description }}",
            "{% endfor %}");
        return Verify(result);
    }
}
