public class EditableFieldGeneratorTests
{
    const string editableModel =
        """
        using System;
        using Parchment;

        namespace Sample;

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

        [ParchmentModel("template.docx")]
        public partial class Order
        {
            public string Number { get; set; } = "";

            [EditableField]
            public string PurchaseOrder { get; set; } = "";

            [EditableField]
            public bool Approved { get; set; }

            [EditableField(DateFormat = "dd MMM yyyy")]
            public DateTime Delivery { get; set; }

            [EditableField]
            public DateTimeOffset SignedAt { get; set; }

            [EditableField]
            public TimeOnly PickupTime { get; set; }

            [EditableField]
            public Status Status { get; set; }

            [EditableField]
            public decimal? Discount { get; set; }

            [EditableField(MultiLine = true)]
            public string Instructions { get; set; } = "";

            public Customer Customer { get; set; } = new();
        }
        """;

    [Test]
    public Task AllKinds_EmitsEditableEntries()
    {
        // Covers per-kind mapping, nullable value type (decimal?), nullable annotated reference
        // (Customer.Email), MultiLine / DateFormat options, and the parent-routed setter +
        // CanReach emission for the nested path.
        var result = GeneratorDriver.Run(
            editableModel,
            "{{ PurchaseOrder }}",
            "{{ Approved }}",
            "{{ Delivery }}",
            "{{ SignedAt }}",
            "{{ PickupTime }}",
            "{{ Status }}",
            "{{ Discount }}",
            "{{ Instructions }}",
            "{{ Customer.Email }}");
        return Verify(result);
    }

    [Test]
    public Task HtmlEditable_EmitsHtmlKind()
    {
        // [Html] + [EditableField] emits an editable entry of kind Html (not a read-only format
        // entry) with the normal string getter/setter.
        var source =
            """
            using Parchment;

            namespace Sample;

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class HtmlAttribute : System.Attribute
            {
            }

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                [EditableField]
                [Html]
                public string Body { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Body }}");
        return Verify(result);
    }

    [Test]
    public Task UnsupportedType_ReportsParch013()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                [EditableField]
                public List<string> Items { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        return Verify(result);
    }

    [Test]
    public async Task NullableBool_ReportsParch013()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                [EditableField]
                public bool? Maybe { get; set; }
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH013");
    }

    [Test]
    public async Task NoUsableSetter_ReportsParch014()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                public string Prefix = "";

                [EditableField]
                public string Computed => Prefix;

                [EditableField]
                public string InitOnly { get; init; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Count(_ => _ == "PARCH014")).IsEqualTo(2);
    }

    [Test]
    public async Task ConflictingAttribute_ReportsParch015()
    {
        // [Markdown] + [EditableField] stays a conflict (editable rich text is HTML-only).
        var source =
            """
            using Parchment;

            namespace Sample;

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class MarkdownAttribute : System.Attribute
            {
            }

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                [EditableField]
                [Markdown]
                public string Body { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH015");
    }

    [Test]
    public async Task HtmlEditable_IsAllowed()
    {
        // [Html] + [EditableField] is supported — an editable rich-content block, no PARCH015.
        var source =
            """
            using Parchment;

            namespace Sample;

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class HtmlAttribute : System.Attribute
            {
            }

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                [EditableField]
                [Html]
                public string Body { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Body }}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Contains("PARCH015")).IsFalse();
        await Assert.That(ids.Any(_ => _.StartsWith("PARCH"))).IsFalse();
    }

    [Test]
    public async Task TokenWithFilter_ReportsParch016()
    {
        var result = GeneratorDriver.Run(editableModel, "{{ PurchaseOrder | upcase }}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH016");
    }

    [Test]
    public async Task DuplicateBodyToken_ReportsParch017()
    {
        var result = GeneratorDriver.Run(
            editableModel,
            "{{ PurchaseOrder }}",
            "again: {{ PurchaseOrder }}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Count(_ => _ == "PARCH017")).IsEqualTo(1);
    }

    [Test]
    public async Task HeaderMirror_DoesNotReportDuplicate()
    {
        // The runtime dispatches editable fields only in the document body; a header occurrence
        // renders as a plain read-only mirror. Duplicate detection must be body-scoped.
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            editableModel,
            ("template.docx", GeneratorDriver.BuildDocxBytesWithHeader(
                bodyParagraphs: ["{{ PurchaseOrder }}"],
                headerParagraphs: ["PO: {{ PurchaseOrder }}"])));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEmpty();
    }

    [Test]
    public async Task RootPathTokenInsideLoop_ReportsParch018Warning()
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
            public partial class Order
            {
                [EditableField]
                public string PurchaseOrder { get; set; } = "";

                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "{{ PurchaseOrder }}",
            "{% endfor %}");
        var diagnostics = result.Results.Single().Diagnostics;
        var warning = diagnostics.Single(_ => _.Id == "PARCH018");
        await Assert.That(warning.Severity).IsEqualTo(DiagnosticSeverity.Warning);
    }

    [Test]
    public async Task LoopVariablePathToken_ReportsParch018Warning()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                [EditableField]
                public string Note { get; set; } = "";
            }

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "{{ line.Note }}",
            "{% endfor %}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH018");
    }

    [Test]
    public async Task StaticEditableMember_IsIgnored()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx")]
            public partial class Order
            {
                [EditableField]
                public static string Logo { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Logo }}");
        var runResult = result.Results.Single();
        await Assert.That(runResult.Diagnostics).IsEmpty();

        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).DoesNotContain("_Editables");
    }

    [Test]
    public async Task ProtectionNone_IsPassedToRegistration()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.docx", Protection = ProtectionMode.None)]
            public partial class Order
            {
                [EditableField]
                public string PurchaseOrder { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ PurchaseOrder }}");
        var generated = result.Results.Single().GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).Contains("store.RegisterDocxTemplate<global::Sample.Order>(TemplateName, path, global::Parchment.ProtectionMode.None);");
    }

    [Test]
    public async Task MarkdownTarget_IgnoresEditableMembers()
    {
        // Editable fields are docx-only, mirroring [Html]/[ExcelsiorTable]: the runtime markdown
        // flow never builds the editable map, so the SG stays silent too.
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel("template.md")]
            public partial class Order
            {
                [EditableField]
                public string PurchaseOrder { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.RunMarkdown(source, "PO: {{ PurchaseOrder | upcase }}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEmpty();
    }
}
