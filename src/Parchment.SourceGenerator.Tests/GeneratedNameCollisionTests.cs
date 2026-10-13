/// <summary>
/// Names the generator invents, or emits unqualified, must not meet a legal name in the consumer's
/// code. Each case compiles the generated output beside the model and expects no warning or error:
/// a hiding warning (CS0108) is a build break under <c>TreatWarningsAsErrors</c>.
/// </summary>
public class GeneratedNameCollisionTests
{
    // Stands in for the parts of Parchment, Fluid and Excelsior the generated code calls, with the
    // real signatures, so the output binds without the test project referencing the runtime.
    const string runtimeStubs =
        """
        namespace Fluid
        {
            public interface IMemberAccessor { }
        }

        namespace Fluid.Accessors
        {
            public sealed class DelegateAccessor(System.Func<object, string, object?> getter) : Fluid.IMemberAccessor
            {
                public System.Func<object, string, object?> Getter { get; } = getter;
            }
        }

        namespace Excelsior
        {
            public sealed class WordTableBuilder<T> { }
        }

        namespace Parchment.Generated
        {
            using System;
            using System.Collections.Generic;

            public enum FormatMapKind { Html, Markdown }

            public enum EditableFieldKind { Text, Number, Checkbox, Date, DateTimeOffset, Time, DropDown, Html }

            public sealed record ExcelsiorTableMapEntry(
                string DottedPath,
                Type ElementType,
                Func<object, object?> Getter,
                string? HeadingParagraphStyle = null,
                string? BodyParagraphStyle = null,
                string? TableStyle = null,
                Action<object>? Configure = null);

            public sealed record FormatMapEntry(string DottedPath, FormatMapKind Format, Func<object, object?> Getter);

            public sealed record StringListMapEntry(string DottedPath, Func<object, object?> Getter);

            public sealed record EditableFieldMapEntry(
                string DottedPath,
                EditableFieldKind Kind,
                Type ClrType,
                bool IsNullable,
                Func<object, object?> Getter,
                Action<object, object?> Setter,
                Func<object, bool> CanReach,
                bool MultiLine,
                string? DateFormat);

            public sealed record CollectionFieldMapEntry(
                string DottedPath,
                Type ElementType,
                Action<object, object?> Setter,
                Func<object, bool> CanReach,
                Func<object> ElementFactory,
                IReadOnlyList<EditableFieldMapEntry> ElementFields,
                bool IsArray);

            [AttributeUsage(AttributeTargets.Class, Inherited = false)]
            public sealed class EmbeddedBytesAttribute(string base64) : Attribute
            {
                public string Base64 { get; } = base64;
            }

            public static class GeneratedRegistration
            {
                public static void RegisterDocxTemplate(Type modelType, Type embedded, global::Parchment.ProtectionMode protection = global::Parchment.ProtectionMode.WhenEditable) { }
                public static void RegisterMarkdownTemplate(Type modelType, string markdown, byte[]? styleSource = null) { }
                public static void RegisterMarkdownTemplate(Type modelType, string markdown, Type embeddedStyle) { }
                public static void RegisterFluidAccessors(Type type, IEnumerable<KeyValuePair<string, global::Fluid.IMemberAccessor>> accessors) { }
                public static void RegisterExcelsiorTable(Type modelType, IEnumerable<ExcelsiorTableMapEntry> entries) { }
                public static void RegisterFormat(Type modelType, IEnumerable<FormatMapEntry> entries) { }
                public static void RegisterStringList(Type modelType, IEnumerable<StringListMapEntry> entries) { }
                public static void RegisterEditable(Type modelType, IEnumerable<EditableFieldMapEntry> entries) { }
                public static void RegisterEditableCollections(Type modelType, IEnumerable<CollectionFieldMapEntry> entries) { }
            }
        }
        """;

    static (List<string> Diagnostics, string Generated) Compile(string source, params TemplateFile[] files)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(source, files);
        var compilation = setup.Compilation.AddSyntaxTrees(
            CSharpSyntaxTree.ParseText(runtimeStubs, (CSharpParseOptions) setup.Compilation.SyntaxTrees[0].Options));
        var driver = setup.Driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out var driverDiagnostics);
        var diagnostics = driverDiagnostics
            .Concat(driver.GetRunResult().Diagnostics)
            .Concat(output.GetDiagnostics())
            .Where(_ => _.Severity >= DiagnosticSeverity.Warning)
            .Select(_ => _.ToString())
            .ToList();
        var generated = string.Concat(driver.GetRunResult().GeneratedTrees.Select(_ => _.GetText().ToString()));
        return (diagnostics, generated);
    }

    static TemplateFile Markdown(string typeName, string markdown) =>
        new($"{typeName}.parchment.md", Encoding.UTF8.GetBytes(markdown));

    // `A_B` and `A.B` once flattened to one field name (CS0102).
    [Test]
    public async Task UnderscoreAndNestedTypeNames_GetDistinctFields()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public class A_B { public string X { get; set; } = ""; }

            public class A { public class B { public string Y { get; set; } = ""; } }

            [ParchmentModel]
            public partial class M
            {
                public A_B One { get; set; } = new();
                public A.B Two { get; set; } = new();
            }
            """;
        var (diagnostics, generated) = Compile(source, Markdown("M", "{{ One.X }} {{ Two.Y }}"));

        await Assert.That(diagnostics).IsEmpty();
        await Assert.That(generated).Contains("_Accessors_global__Sample_A_95_B =");
        await Assert.That(generated).Contains("_Accessors_global__Sample_A_B =");
    }

    [Test]
    public async Task ToIdentifier_IsInjective()
    {
        string[] names =
        [
            "global::Sample.A_B",
            "global::Sample.A.B",
            "global::Sample.A._B",
            "global::Sample.A_.B",
            "global::Sample.A__B",
            "global::Sample_A.B",
            "global::Sample.A95.B",
            "global::Sample.A.@class",
            "global::Sample.A._64_class",
            "global::Sample.Box<global::Sample.A>",
            "global::Sample.Box<global::Sample.A>[]",
            "global::Sample.Box_60_global__Sample_A_62_",
            "global::Sample.Box<global::Sample.A, global::Sample.B>",
            "global::Sample.A?",
            "(int, string)"
        ];
        var identifiers = names.Select(AccessorEmission.ToIdentifier).ToList();

        await Assert.That(identifiers.Distinct(StringComparer.Ordinal).Count()).IsEqualTo(names.Length);
        foreach (var identifier in identifiers)
        {
            await Assert.That(SyntaxFacts.IsValidIdentifier("_" + identifier)).IsTrue();
        }

        // A name with no underscore keeps the spelling it always had.
        await Assert.That(AccessorEmission.ToIdentifier("global::Sample.Invoice")).IsEqualTo("global__Sample_Invoice");
    }

    // The accessor table was the one framework type emitted without `global::`, resolved through a
    // using directive that a type in the model's own namespace takes precedence over (CS1729).
    [Test]
    public async Task UserTypesNamedLikeEmittedTypes_DoNotCapture()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public class KeyValuePair<TKey, TValue> { }

            public class Type { }

            public class Convert { }

            public class Activator { }

            public class Line
            {
                [EditableField]
                public string Text { get; set; } = "";
            }

            [ParchmentModel]
            public partial class M
            {
                public string Name { get; set; } = "";

                public string System { get; set; } = "";

                public string Parchment { get; set; } = "";

                public string Fluid { get; set; } = "";

                public string Excelsior { get; set; } = "";

                [EditableField]
                public global::System.Collections.Generic.List<Line> Lines { get; set; } = new();

                [global::Parchment.ExcelsiorTable]
                public global::System.Collections.Generic.List<Line> Table { get; set; } = new();
            }
            """;
        var (diagnostics, generated) = Compile(source, Markdown("M", "{{ Name }} {{ System }}"));

        await Assert.That(diagnostics).IsEmpty();
        await Assert.That(generated).Contains("global::System.Collections.Generic.KeyValuePair<string, global::Fluid.IMemberAccessor>[]");
        await Assert.That(generated).DoesNotContain("using ");
    }

    // User namespaces that end in the names the generated code roots its own references at.
    [Test]
    public async Task UserNamespacesNamedLikeEmittedNamespaces_DoNotCapture()
    {
        var source =
            """
            namespace Sample.System { public class Thing { public string N { get; set; } = ""; } }
            namespace Sample.Parchment { public class Thing { public string N { get; set; } = ""; } }
            namespace Sample.Fluid { public class Thing { public string N { get; set; } = ""; } }
            namespace Sample.Excelsior { public class Thing { public string N { get; set; } = ""; } }
            namespace Sample
            {
                [global::Parchment.ParchmentModel]
                public partial class M
                {
                    public Sample.System.Thing A { get; set; } = new();
                    public Sample.Parchment.Thing B { get; set; } = new();
                    public Sample.Fluid.Thing C { get; set; } = new();
                    [global::Parchment.ExcelsiorTable] public global::System.Collections.Generic.List<Sample.Excelsior.Thing> D { get; set; } = new();
                    [global::Parchment.EditableField] public string E { get; set; } = "";
                }
            }
            """;
        var (diagnostics, _) = Compile(source, Markdown("M", "{{ A.N }}"));

        await Assert.That(diagnostics).IsEmpty();
    }

    // A derived model once hid the nested type of a base with the same simple name (CS0108).
    [Test]
    public async Task DerivedModelWithItsBasesSimpleName_HidesNothing()
    {
        var source =
            """
            using Parchment;

            namespace Contracts
            {
                [ParchmentBindable]
                public partial class Invoice { public string Number { get; set; } = ""; }

                [ParchmentModel]
                public partial class Letter { public string Name { get; set; } = ""; }
            }

            namespace Reports
            {
                [ParchmentBindable]
                public partial class Invoice : Contracts.Invoice { public string Extra { get; set; } = ""; }

                public partial class Outer
                {
                    [ParchmentBindable]
                    public partial class Letter : Contracts.Letter { public string Extra { get; set; } = ""; }
                }
            }
            """;
        var (diagnostics, generated) = Compile(source, Markdown("Letter", "{{ Name }}"));

        await Assert.That(diagnostics).IsEmpty();
        await Assert.That(generated).Contains("internal static class ParchmentGenerated_Contracts_Invoice");
        await Assert.That(generated).Contains("internal static class ParchmentGenerated_Reports_Invoice");
        await Assert.That(generated).Contains("internal static class ParchmentGenerated_Reports_Outer_Letter");
    }
}
