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

        [ParchmentModel]
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

            [ParchmentModel]
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

            [ParchmentModel]
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

            [ParchmentModel]
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
    public async Task FlagsEnum_ReportsParch013()
    {
        // A dropdown holds one item, so a combined value (Read | Write) has no item to select.
        // The nullable member and the member of a collection element take the same path.
        var source =
            """
            using System;
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            [Flags]
            public enum Access
            {
                None = 0,
                Read = 1,
                Write = 2
            }

            public class Grant
            {
                [EditableField]
                public Access Access { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public Access Access { get; set; }

                [EditableField]
                public Access? Optional { get; set; }

                [EditableField]
                public List<Grant> Grants { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var diagnostics = result.Results.Single().Diagnostics.Where(_ => _.Id == "PARCH013").ToList();
        await Assert.That(diagnostics.Count).IsEqualTo(3);
        await Assert.That(diagnostics[0].GetMessage()).Contains("[Flags]");
    }

    [Test]
    public async Task EnumWithoutFlags_IsADropDown()
    {
        // Only the attribute decides: an enum that names a combination without [Flags] still
        // offers each named constant as one dropdown item.
        var source =
            """
            using Parchment;

            namespace Sample;

            public enum Level
            {
                Low = 1,
                High = 2,
                Both = Low | High
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public Level Level { get; set; }

                [EditableField]
                public Level? Optional { get; set; }
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Level }} {{ Optional }}");
        await Assert.That(ParchIds(result)).IsEmpty();
        var generated = result.Results.Single().GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).Contains("EditableFieldKind.DropDown");
    }

    [Test]
    public async Task CollectionElementWithoutEditableMembers_ReportsParch022()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class EmptyElement
            {
                public string Name { get; set; } = "";
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public List<EmptyElement> Items { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Select(_ => _.Id)).Contains("PARCH022");
        await Assert.That(diagnostics.Single(_ => _.Id == "PARCH022").GetMessage()).Contains("no [EditableField] members");
    }

    [Test]
    public async Task CollectionElementWithoutParameterlessCtor_ReportsParch022()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Budget
            {
                public Budget(string year) => Year = year;

                [EditableField]
                public string Year { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public List<Budget> Items { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Select(_ => _.Id)).Contains("PARCH022");
        await Assert.That(diagnostics.Single(_ => _.Id == "PARCH022").GetMessage()).Contains("parameterless constructor");
    }

    [Test]
    public async Task NestedEditableCollection_ReportsParch022()
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

            public class Budget
            {
                [EditableField]
                public string Year { get; set; } = "";

                [EditableField]
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public List<Budget> Items { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var diagnostics = result.Results.Single().Diagnostics;
        await Assert.That(diagnostics.Any(_ => _.Id == "PARCH022" && _.GetMessage().Contains("nested"))).IsTrue();
    }

    // Extraction rebuilds an editable collection as a List<T> (a T[] for an array member) and the
    // generated setter casts it to the declared type, so a declared type that cannot hold one
    // used to compile, render, extract, and then throw InvalidCastException from ApplyTo.
    [Test]
    [Arguments("HashSet<Line>")]
    [Arguments("SortedSet<Line>")]
    [Arguments("LinkedList<Line>")]
    [Arguments("Collection<Line>")]
    [Arguments("ObservableCollection<Line>")]
    [Arguments("ImmutableArray<Line>")]
    [Arguments("ImmutableList<Line>")]
    [Arguments("IImmutableList<Line>")]
    [Arguments("ISet<Line>")]
    [Arguments("LineList")]
    [Arguments("ILineList")]
    [Arguments("Line[,]")]
    public async Task CollectionTypeThatCannotHoldTheRebuiltList_ReportsParch022(string type)
    {
        var result = GeneratorDriver.Run(CollectionTypeModel(type), "x");
        var runResult = result.Results.Single();
        var diagnostic = runResult.Diagnostics.Single(_ => _.Id.StartsWith("PARCH", StringComparison.Ordinal));
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH022");
        await Assert.That(diagnostic.GetMessage()).Contains("cannot hold the rebuilt list");

        // No entry is emitted, so the setter whose cast would fail does not exist.
        var generated = string.Concat(runResult.GeneratedSources.Select(_ => _.SourceText.ToString()));
        await Assert.That(generated).DoesNotContain("_EditableCollections");
    }

    [Test]
    [Arguments("List<Line>")]
    [Arguments("Line[]")]
    [Arguments("IList<Line>")]
    [Arguments("ICollection<Line>")]
    [Arguments("IEnumerable<Line>")]
    [Arguments("IReadOnlyList<Line>")]
    [Arguments("IReadOnlyCollection<Line>")]
    [Arguments("List<Line>?")]
    public async Task CollectionTypeThatHoldsTheRebuiltList_EmitsCollectionEntry(string type)
    {
        var result = GeneratorDriver.Run(CollectionTypeModel(type), "x");
        var runResult = result.Results.Single();
        var ids = runResult.Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Any(_ => _.StartsWith("PARCH", StringComparison.Ordinal))).IsFalse();

        var generated = string.Concat(runResult.GeneratedSources.Select(_ => _.SourceText.ToString()));
        await Assert.That(generated).Contains("_EditableCollections");
    }

    // A setter through a struct assigns to a copy, so nothing the generator could emit writes the
    // value back. Reported at the member; what would not compile is not emitted.
    [Test]
    [Arguments("struct")]
    [Arguments("record struct")]
    public async Task EditableMemberInStruct_ReportsParch032(string kind)
    {
        var source =
            $$"""
            using Parchment;

            namespace Sample;

            public {{kind}} Address
            {
                [EditableField]
                public string City { get; set; }

                [EditableField]
                public int Floor;

                public string Country { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public string Number { get; set; } = "";

                public Address Address { get; set; }
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Address.Country }}");
        var runResult = result.Results.Single();
        var messages = runResult.Diagnostics
            .Where(_ => _.Id.StartsWith("PARCH", StringComparison.Ordinal))
            .Select(_ => $"{_.Id}: {_.GetMessage()}")
            .ToList();
        await Assert.That(messages.Count).IsEqualTo(2);
        await Assert.That(messages.All(_ => _.StartsWith("PARCH032", StringComparison.Ordinal))).IsTrue();
        await Assert.That(messages.Any(_ => _.Contains("'Sample.Address.City'") && _.Contains("the struct 'Sample.Address'"))).IsTrue();
        await Assert.That(messages.Any(_ => _.Contains("'Sample.Address.Floor'"))).IsTrue();

        // The member beside it on the class keeps its entry, and the struct's members still bind.
        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).Contains("new(\"Number\", ");
        await Assert.That(generated).DoesNotContain("\"Address.City\"");
        await Assert.That(generated).DoesNotContain("\"Address.Floor\"");
        await Assert.That(generated).Contains("((global::Sample.Address)o).City");
    }

    [Test]
    public async Task EditableMemberInClassBelowStruct_ReportsParch032()
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

            public class Contact
            {
                [EditableField]
                public string Email { get; set; } = "";

                [EditableField]
                public List<Line> Lines { get; set; } = new();
            }

            public struct Party
            {
                public Contact Contact { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                public Party Buyer { get; set; }

                public Party Seller { get; set; }
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var runResult = result.Results.Single();
        var messages = runResult.Diagnostics
            .Where(_ => _.Id.StartsWith("PARCH", StringComparison.Ordinal))
            .Select(_ => $"{_.Id}: {_.GetMessage()}")
            .ToList();

        // Once per member, though two paths reach each.
        await Assert.That(messages.Count).IsEqualTo(2);
        await Assert.That(messages.Any(_ => _.StartsWith("PARCH032", StringComparison.Ordinal) && _.Contains("'Sample.Contact.Email'") && _.Contains("the struct 'Sample.Party'"))).IsTrue();
        await Assert.That(messages.Any(_ => _.StartsWith("PARCH032", StringComparison.Ordinal) && _.Contains("'Sample.Contact.Lines'"))).IsTrue();

        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).DoesNotContain("_Editables");
        await Assert.That(generated).DoesNotContain("_EditableCollections");
    }

    [Test]
    [Arguments("List<Row>")]
    [Arguments("Row[]")]
    public async Task StructCollectionElement_ReportsParch022(string collection)
    {
        var source =
            $$"""
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public struct Row
            {
                [EditableField]
                public int Year { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public {{collection}} Rows { get; set; } = [];
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for row in Rows %}",
            "{{ row.Year }}",
            "{% endfor %}");
        var runResult = result.Results.Single();
        var errors = runResult.Diagnostics
            .Where(_ => _.Severity == DiagnosticSeverity.Error)
            .ToList();
        await Assert.That(errors.Count).IsEqualTo(1);
        await Assert.That(errors[0].Id).IsEqualTo("PARCH022");
        await Assert.That(errors[0].GetMessage()).Contains("'Sample.Order.Rows'");
        await Assert.That(errors[0].GetMessage()).Contains("'Sample.Row' is a struct");

        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).DoesNotContain("_EditableCollections");
    }

    [Test]
    public async Task EditableMemberInStructInsideCollectionElement_ReportsParch032()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public struct Money
            {
                [EditableField]
                public decimal Amount { get; set; }
            }

            public class Line
            {
                [EditableField]
                public string Note { get; set; } = "";

                public Money Price { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for line in Lines %}",
            "{{ line.Note }}",
            "{% endfor %}");
        var runResult = result.Results.Single();
        var errors = runResult.Diagnostics
            .Where(_ => _.Severity == DiagnosticSeverity.Error)
            .ToList();
        await Assert.That(errors.Count).IsEqualTo(1);
        await Assert.That(errors[0].Id).IsEqualTo("PARCH032");
        await Assert.That(errors[0].GetMessage()).Contains("'Sample.Money.Amount'");

        // The collection and its other member are still emitted.
        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).Contains("_EditableCollections");
        await Assert.That(generated).Contains("new(\"Note\", ");
        await Assert.That(generated).DoesNotContain("\"Price.Amount\"");
    }

    [Test]
    public async Task StructMembersWithoutEditableField_NoDiagnostics()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public struct Address
            {
                public string City { get; set; }
            }

            public struct Row
            {
                public int Year { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                public Address Address { get; set; }

                public List<Row> Rows { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{{ Address.City }}",
            "{% for row in Rows %}",
            "{{ row.Year }}",
            "{% endfor %}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Any(_ => _.StartsWith("PARCH", StringComparison.Ordinal))).IsFalse();
    }

    static string CollectionTypeModel(string type) =>
        $$"""
          #nullable enable
          using System.Collections.Generic;
          using System.Collections.Immutable;
          using System.Collections.ObjectModel;
          using Parchment;

          namespace Sample;

          public class Line
          {
              [EditableField]
              public string Note { get; set; } = "";
          }

          public class LineList : List<Line>;

          public interface ILineList : IList<Line>;

          [ParchmentModel]
          public partial class Order
          {
              [EditableField]
              public {{type}} Lines { get; set; } = default!;
          }
          """;

    [Test]
    public async Task ConflictingFormatMarkers_ReportsParch023()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class HtmlAttribute : System.Attribute { }

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class MarkdownAttribute : System.Attribute { }

            [ParchmentModel]
            public partial class Order
            {
                [Html]
                [Markdown]
                public string Body { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH023");
    }

    [Test]
    public async Task HtmlContradictedByStringSyntax_ReportsParch023()
    {
        var source =
            """
            using System.Diagnostics.CodeAnalysis;
            using Parchment;

            namespace Sample;

            [System.AttributeUsage(System.AttributeTargets.Property)]
            public sealed class HtmlAttribute : System.Attribute { }

            [ParchmentModel]
            public partial class Order
            {
                [Html]
                [StringSyntax("markdown")]
                public string Body { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "x");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH023");
    }

    [Test]
    public async Task NoUsableSetter_ReportsParch014()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
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

            [ParchmentModel]
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

            [ParchmentModel]
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

    static List<Diagnostic> BlockDiagnostics(string model, params OpenXmlElement[] blocks)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            model,
            new TemplateFile($"{GeneratorDriver.ModelTypeName(model)}.parchment.docx", GeneratorDriver.BuildDocxBytesFromBlocks(blocks)));
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult().Results.Single().Diagnostics.ToList();
    }

    static Paragraph ControlParagraph(OpenXmlElement? marker, string text) =>
        new(
            new Run(new Text("before ")
            {
                Space = SpaceProcessingModeValues.Preserve
            }),
            new SdtRun(
                new SdtProperties(marker == null ? [] : [marker]),
                new SdtContentRun(new Run(new Text(text)))));

    public static IEnumerable<Func<OpenXmlElement>> ValueControlMarkers()
    {
        yield return () => new SdtContentText();
        yield return () => new SdtContentDate();
        yield return () => new SdtContentDropDownList();
        yield return () => new SdtContentComboBox();
        yield return () => new DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox();
        yield return () => new SdtContentPicture();
    }

    [Test]
    [MethodDataSource(nameof(ValueControlMarkers))]
    public async Task TokenInsideValueControl_ReportsParch016(OpenXmlElement marker)
    {
        // Lockstep with the runtime's EditableTokenValidator: the field would be a content
        // control nested in one Word does not let hold another.
        var diagnostics = BlockDiagnostics(editableModel, ControlParagraph(marker, "{{ PurchaseOrder }}"));
        var diagnostic = diagnostics.Single();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH016");
        await Assert.That(diagnostic.GetMessage()).Contains("content control");
    }

    [Test]
    public async Task TokenInsideBlockLevelValueControl_ReportsParch016()
    {
        var diagnostics = BlockDiagnostics(
            editableModel,
            new SdtBlock(
                new SdtProperties(new SdtContentText()),
                new SdtContentBlock(GeneratorDriver.BuildParagraph("{{ PurchaseOrder }}"))));
        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH016"]);
    }

    [Test]
    public async Task TokenInsideRichTextControl_NoDiagnostics()
    {
        var diagnostics = BlockDiagnostics(editableModel, ControlParagraph(null, "{{ PurchaseOrder }}"));
        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task TokenBesideValueControl_NoDiagnostics()
    {
        var paragraph = ControlParagraph(new SdtContentText(), "static");
        paragraph.AppendChild(new Run(new Text("{{ PurchaseOrder }}")));
        var diagnostics = BlockDiagnostics(editableModel, paragraph);
        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task CollectionItemTokenInsideValueControl_ReportsParch016()
    {
        var diagnostics = BlockDiagnostics(
            collectionModel,
            GeneratorDriver.BuildParagraph("{% for line in Lines %}"),
            ControlParagraph(new SdtContentText(), "{{ line.Note }}"),
            GeneratorDriver.BuildParagraph("{% endfor %}"));
        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH016"]);
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
            new TemplateFile($"{GeneratorDriver.ModelTypeName(editableModel)}.parchment.docx", GeneratorDriver.BuildDocxBytesWithHeader(
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

            [ParchmentModel]
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

    // A range loop variable shadows a model member of the same name, as it does at render, so the
    // token is the loop value and not the editable field.
    [Test]
    public async Task RangeLoopVariable_ShadowingEditableMember_NoParch018()
    {
        var result = GeneratorDriver.Run(
            editableModel,
            "{% for PurchaseOrder in (1..3) %}",
            "{{ PurchaseOrder }}",
            "{% endfor %}");
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).IsEmpty();
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

            [ParchmentModel]
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

    const string collectionModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            [EditableField]
            public string Note { get; set; } = "";

            [EditableField]
            public decimal Amount { get; set; }
        }

        [ParchmentModel]
        public partial class Order
        {
            [EditableField]
            public List<Line> Lines { get; set; } = new();
        }
        """;

    [Test]
    public Task EditableCollection_EmitsCollectionEntry()
    {
        // [EditableField] on a POCO collection → a repeating-section registration
        // (RegisterEditableCollections) with element-rooted sub-entry getters/setters; the
        // loop-variable tokens do NOT report PARCH018.
        var result = GeneratorDriver.Run(
            collectionModel,
            "{% for line in Lines %}",
            "{{ line.Note }} {{ line.Amount }}",
            "{% endfor %}");
        return Verify(result);
    }

    [Test]
    public async Task EditableCollection_NestedElementMember_SetterCreatesParents()
    {
        // Extraction builds each item from nothing, so the element-rooted setter creates the
        // objects on the way to a nested member - but only the ones it can: an abstract type or a
        // get-only member is left to the null check, and the root-level setter is unchanged.
        var result = GeneratorDriver.Run(
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Depot
            {
                [EditableField]
                public string? Code { get; set; }
            }

            public abstract class Carrier
            {
                [EditableField]
                public string? Contact { get; set; }
            }

            public class Shipping
            {
                [EditableField]
                public int Boxes { get; set; }

                public Depot? Depot { get; set; }

                public Depot Fixed { get; } = new();
            }

            public class Line
            {
                [EditableField]
                public string? Name { get; set; }

                public Shipping? Shipping { get; set; }

                public Carrier? Carrier { get; set; }
            }

            [ParchmentModel]
            public partial class Order
            {
                public Shipping? Shipping { get; set; }

                [EditableField]
                public List<Line> Lines { get; set; } = new();
            }
            """,
            "{% for line in Lines %}",
            "{{ line.Shipping.Boxes }} {{ line.Shipping.Depot.Code }} {{ line.Shipping.Fixed.Code }} {{ line.Carrier.Contact }}",
            "{% endfor %}");
        var runResult = result.Results.Single();
        var diagnostics = runResult.Diagnostics.Select(_ => _.ToString()).ToList();
        await Assert.That(diagnostics).IsEmpty();

        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        const string createShipping = "(o, v) => { var p0 = (global::Sample.Line)o; var p1 = p0.Shipping; if (p0 != null && p1 == null) { p1 = (global::Sample.Shipping)global::System.Activator.CreateInstance(typeof(global::Sample.Shipping))!; p0.Shipping = p1; } ";
        await Assert.That(generated).Contains(createShipping + "if (p1 != null) { p1.Boxes = (int)v!; } }");
        await Assert.That(generated).Contains(createShipping + "var p2 = p1?.Depot; if (p1 != null && p2 == null) { p2 = (global::Sample.Depot)global::System.Activator.CreateInstance(typeof(global::Sample.Depot))!; p1.Depot = p2; } if (p2 != null) { p2.Code = (string?)v; } }");
        await Assert.That(generated).Contains(createShipping + "var p2 = p1?.Fixed; if (p2 != null) { p2.Code = (string?)v; } }");
        await Assert.That(generated).Contains("(o, v) => { var p = ((global::Sample.Line)o).Carrier; if (p != null) { p.Contact = (string?)v; } }");
        await Assert.That(generated).Contains("(o, v) => { var p = ((global::Sample.Order)o).Shipping; if (p != null) { p.Boxes = (int)v!; } }");
    }

    [Test]
    public async Task EditableCollection_NoDiagnostics()
    {
        var result = GeneratorDriver.Run(
            collectionModel,
            "{% for line in Lines %}",
            "{{ line.Note }} {{ line.Amount }}",
            "{% endfor %}");
        var runResult = result.Results.Single();
        var ids = runResult.Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids.Any(_ => _.StartsWith("PARCH", StringComparison.Ordinal))).IsFalse();

        var generated = runResult.GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).Contains("_EditableCollections");
        await Assert.That(generated).Contains("RegisterEditableCollections");
    }

    [Test]
    public async Task StaticEditableMember_IsIgnored()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public static string Logo { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Logo }}");
        var runResult = result.Results.Single();

        // Ignored, but no longer silently: the attribute does nothing on a static member, so the
        // generator says so rather than leaving it to be discovered at render.
        var warning = runResult.Diagnostics.SingleOrDefault(_ => _.Id == "PARCH019");
        await Assert.That(warning).IsNotNull();
        await Assert.That(warning!.GetMessage()).Contains("EditableField");

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

            [ParchmentModel(Protection = ProtectionMode.None)]
            public partial class Order
            {
                [EditableField]
                public string PurchaseOrder { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ PurchaseOrder }}");
        var generated = result.Results.Single().GeneratedSources.Single().SourceText.ToString();
        await Assert.That(generated).Contains("global::Parchment.ProtectionMode.None);");
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

            [ParchmentModel]
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

    static List<string> ParchIds(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .Where(_ => _.StartsWith("PARCH", StringComparison.Ordinal))
            .ToList();

    // Extraction replaces the collection with the rows in the document, so a loop that rendered a
    // window of it, or rendered it backwards, would have that written back.
    [Test]
    [Arguments("{% for line in Lines limit:1 %}")]
    [Arguments("{% for line in Lines offset:1 %}")]
    [Arguments("{% for line in Lines reversed %}")]
    public async Task EditableCollection_LoopModifier_ReportsParch027(string openTag)
    {
        var result = GeneratorDriver.Run(
            collectionModel,
            openTag,
            "{{ line.Note }} {{ line.Amount }}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH027"]);
    }

    // A nested loop runs with an empty editable map and renders read-only, so nothing it shows is
    // extracted and a modifier there is harmless. PARCH018 is the existing warning for that nesting.
    [Test]
    public async Task EditableCollection_LoopModifier_InsideAnotherLoop_NoParch027()
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

            [ParchmentModel]
            public partial class Order
            {
                public List<string> Pages { get; set; } = new();

                [EditableField]
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{% for page in Pages %}",
            "{% for line in Lines limit:1 %}",
            "{{ line.Note }}",
            "{% endfor %}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).DoesNotContain("PARCH027");
    }

    // An item member reached through the loop variable is a control like a root member: the
    // runtime builds it from the member, so a filter would be dropped.
    [Test]
    public async Task EditableCollection_ItemMemberWithFilter_ReportsParch016()
    {
        var result = GeneratorDriver.Run(
            collectionModel,
            "{% for line in Lines %}",
            "{{ line.Note }} {{ line.Amount | plus: 1 }}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH016"]);
    }

    [Test]
    public async Task EditableCollection_ItemMemberBoundTwice_ReportsParch017()
    {
        var result = GeneratorDriver.Run(
            collectionModel,
            "{% for line in Lines %}",
            "{{ line.Note }} {{ line.Amount }}",
            "again: {{ line.Note }}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH017"]);
    }

    const string entryModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            [EditableField]
            public string Note { get; set; } = "";

            public string Code { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Order
        {
            [EditableField]
            public string Note { get; set; } = "";

            [EditableField]
            public List<Line> Lines { get; set; } = new();
        }
        """;

    // Extraction reads one repeating section per collection, and the first control per member in
    // a row, so a second binding of either is an edit that would be discarded.
    [Test]
    public async Task EditableCollection_SecondBindingLoop_ReportsParch017()
    {
        var result = GeneratorDriver.Run(
            entryModel,
            "{% for line in Lines %}",
            "{{ line.Note }}",
            "{% endfor %}",
            "{% for other in Lines %}",
            "{% if other.Code %}",
            "{{ other.Note }}",
            "{% endif %}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH017"]);
    }

    [Test]
    public async Task EditableCollection_MemberTwiceInLoop_ReportsParch017()
    {
        var result = GeneratorDriver.Run(
            entryModel,
            "{% for line in Lines %}",
            "{{ line.Note }} and {{ line.Note }}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH017"]);
    }

    // A loop that binds no control renders read-only, and the root member that shares a row
    // member's name is a different control.
    [Test]
    public async Task EditableCollection_LoopsWithoutControls_NoParch017()
    {
        var result = GeneratorDriver.Run(
            entryModel,
            "{{ Note }}",
            "{% for line in Lines %}",
            "{{ line.Code }}",
            "{% endfor %}",
            "{% if Lines.size > 0 %}",
            "some",
            "{% endif %}",
            "{% for line in Lines %}",
            "{{ line.Note }}",
            "{% endfor %}",
            "{% for line in Lines %}",
            "{% for inner in Lines %}",
            "{{ line.Code }}{{ inner.Code }}",
            "{% endfor %}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).DoesNotContain("PARCH017");
    }

    // A loop nested in the repeating section renders read-only, so a filter or a repeat there is
    // applied and only the read-only warning is due.
    [Test]
    public async Task EditableCollection_ItemMemberInNestedLoop_ReportsOnlyParch018()
    {
        var result = GeneratorDriver.Run(
            collectionModel,
            "{% for line in Lines %}",
            "{{ line.Note }}",
            "{% for other in Lines %}",
            "{{ line.Note | upcase }}",
            "{% endfor %}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH018"]);
    }

    // The collection is not a field itself — the runtime keeps it out of the scalar map — so its
    // own token is an ordinary substitution.
    [Test]
    public async Task EditableCollection_OwnToken_NoDiagnostics()
    {
        var result = GeneratorDriver.Run(
            collectionModel,
            "{{ Lines | size }} of {{ Lines | size }}",
            "{% for line in Lines %}",
            "{{ line.Note }} {{ line.Amount }}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEmpty();
    }

    // Inside any loop a root member renders read-only, whatever the loop iterates: a filter is
    // applied and the body control stays the only one.
    [Test]
    [Arguments("Lines")]
    [Arguments("Pages")]
    public async Task RootTokenInsideLoop_WithFilterBesideBodyControl_ReportsOnlyParch018(string collection)
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

            [ParchmentModel]
            public partial class Order
            {
                [EditableField]
                public string PurchaseOrder { get; set; } = "";

                public List<string> Pages { get; set; } = new();

                [EditableField]
                public List<Line> Lines { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(
            source,
            "{{ PurchaseOrder }}",
            $"{{% for item in {collection} %}}",
            "{{ PurchaseOrder | upcase }}",
            "{% endfor %}");
        await Assert.That(ParchIds(result)).IsEquivalentTo(["PARCH018"]);
    }
}
