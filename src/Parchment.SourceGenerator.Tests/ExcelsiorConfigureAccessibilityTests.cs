/// <summary>
/// <c>[ExcelsiorTable(Configure = ...)]</c> names a method the generated registration calls, and
/// that registration sits in a type nested in the model. The documented form is a private method,
/// which a nested type can see on the model and nowhere else — so each test here compiles the
/// generator's output, against stand-ins for what it references, rather than reading it.
/// </summary>
public class ExcelsiorConfigureAccessibilityTests
{
    const string stubs =
        """
        namespace Fluid
        {
            public interface IMemberAccessor;
        }

        namespace Fluid.Accessors
        {
            public class DelegateAccessor(System.Func<object, object, object?> getter) : Fluid.IMemberAccessor
            {
                public System.Func<object, object, object?> Getter { get; } = getter;
            }
        }

        namespace Excelsior
        {
            public class WordTableBuilder<T>;
        }

        namespace Parchment.Generated
        {
            public sealed record ExcelsiorTableMapEntry(
                string DottedPath,
                System.Type ElementType,
                System.Func<object, object?> Getter,
                string? HeadingParagraphStyle = null,
                string? BodyParagraphStyle = null,
                string? TableStyle = null,
                System.Action<object>? Configure = null);

            public static class GeneratedRegistration
            {
                public static void RegisterFluidAccessors(System.Type type, System.Collections.Generic.KeyValuePair<string, Fluid.IMemberAccessor>[] accessors)
                {
                }

                public static void RegisterExcelsiorTable(System.Type type, ExcelsiorTableMapEntry[] entries)
                {
                }
            }
        }
        """;

    const string prelude =
        """
        using System.Collections.Generic;
        using Excelsior;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Description { get; set; } = "";
        }

        """;

    sealed record Outcome(
        ImmutableArray<Diagnostic> Generator,
        List<string> CompileErrors,
        string Generated);

    static Outcome Run(string source, params MetadataReference[] references)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(prelude + source);
        var compilation = setup.Compilation
            .AddSyntaxTrees(CSharpSyntaxTree.ParseText(stubs))
            .AddReferences(references);
        var driver = setup.Driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out _);
        var result = driver.GetRunResult().Results.Single();
        var errors = output.GetDiagnostics()
            .Where(_ => _.Severity == DiagnosticSeverity.Error)
            .Select(_ => _.ToString())
            .ToList();
        var generated = string.Join('\n', result.GeneratedSources.Select(_ => _.SourceText.ToString()));
        return new(result.Diagnostics, errors, generated);
    }

    static async Task AssertCompiles(Outcome outcome)
    {
        await Assert.That(outcome.Generator).IsEmpty();
        await Assert.That(outcome.CompileErrors).IsEmpty();
    }

    // The diagnostic replaces the call, so it is the only error: nothing generated mentions the
    // method, and what is generated still compiles.
    static async Task AssertReported(Outcome outcome, string method, string type)
    {
        var error = outcome.Generator.Single();
        await Assert.That(error.Id).IsEqualTo("PARCH026");
        await Assert.That(error.GetMessage()).Contains($"'{method}'");
        await Assert.That(error.GetMessage()).Contains($"'{type}'");
        await Assert.That(error.GetMessage()).Contains("partial");
        await Assert.That(outcome.CompileErrors).IsEmpty();
        await Assert.That(outcome.Generated).DoesNotContain(method);
    }

    [Test]
    public async Task PrivateOnRootModel_CalledDirectly()
    {
        var outcome = Run(
            """
            [ParchmentBindable]
            public partial class Invoice
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Invoice.ConfigureLines(");
        await Assert.That(outcome.Generated).DoesNotContain("ParchmentConfigure_");
    }

    [Test]
    public async Task PrivateOnPartialClassReachedThroughModel_Forwarded()
    {
        var outcome = Run(
            """
            public partial class Customer
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ParchmentConfigure_Sample_Invoice_ConfigureLines(");
        await Assert.That(outcome.Generated).Contains(
            """
            namespace Sample
            {
              partial class Customer
              {
                [global::System.ComponentModel.EditorBrowsable(global::System.ComponentModel.EditorBrowsableState.Never)]
                internal static void ParchmentConfigure_Sample_Invoice_ConfigureLines(global::Excelsior.WordTableBuilder<global::Sample.Line> builder) =>
                  ConfigureLines(builder);
              }
            }
            """.Replace("\r\n", "\n"));
    }

    // The bindable type registers itself, calling its own private method directly, and the model
    // reaching it needs the forwarder. Two models reaching one type each get their own.
    [Test]
    public async Task PrivateOnBindableReachedByTwoModels_EachForwarded()
    {
        var outcome = Run(
            """
            [ParchmentBindable]
            public partial class Customer
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }

            [ParchmentBindable]
            public partial class Statement
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ConfigureLines(");
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ParchmentConfigure_Sample_Invoice_ConfigureLines(");
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ParchmentConfigure_Sample_Statement_ConfigureLines(");
    }

    [Test]
    public Task PrivateOnNonPartialClass_Reported()
    {
        var outcome = Run(
            """
            public class Customer
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        return AssertReported(outcome, "ConfigureLines", "Sample.Customer");
    }

    [Test]
    public async Task InternalOnNonPartialClass_CalledDirectly()
    {
        var outcome = Run(
            """
            public class Customer
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                internal static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ConfigureLines(");
    }

    [Test]
    public async Task PrivateOnPartialBaseOfModel_Forwarded()
    {
        var outcome = Run(
            """
            public partial class DocumentBase
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice : DocumentBase
            {
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.DocumentBase.ParchmentConfigure_Sample_Invoice_ConfigureLines(");
    }

    [Test]
    public Task PrivateOnNonPartialBaseOfModel_Reported()
    {
        var outcome = Run(
            """
            public class DocumentBase
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice : DocumentBase
            {
            }
            """);

        return AssertReported(outcome, "ConfigureLines", "Sample.DocumentBase");
    }

    // A protected static is reachable from a type nested in a derived one, so nothing is forwarded.
    [Test]
    public async Task ProtectedOnNonPartialBaseOfModel_CalledDirectly()
    {
        var outcome = Run(
            """
            public class DocumentBase
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                protected static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice : DocumentBase
            {
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.DocumentBase.ConfigureLines(");
    }

    // The partial part restates the type parameters, and the forwarder's parameter is in terms of
    // them; the call site closes both.
    [Test]
    public async Task PrivateOnGenericPartialClass_Forwarded()
    {
        var outcome = Run(
            """
            public partial class Section<T>
            {
                [ExcelsiorTable(Configure = nameof(ConfigureItems))]
                public List<T> Items { get; set; } = new();

                static void ConfigureItems(WordTableBuilder<T> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Section<Line> Section { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("partial class Section<T>");
        await Assert.That(outcome.Generated).Contains("ParchmentConfigure_Sample_Invoice_ConfigureItems(global::Excelsior.WordTableBuilder<T> builder)");
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Section<global::Sample.Line>.ParchmentConfigure_Sample_Invoice_ConfigureItems(");
    }

    [Test]
    public async Task PrivateOnNestedPartialClass_Forwarded()
    {
        var outcome = Run(
            """
            public partial class Crm
            {
                public partial record Customer
                {
                    [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                    public List<Line> Lines { get; set; } = new();

                    static void ConfigureLines(WordTableBuilder<Line> builder)
                    {
                    }
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Crm.Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Crm.Customer.ParchmentConfigure_Sample_Invoice_ConfigureLines(");
    }

    // Every enclosing type has to be partial for the part to be declared at all.
    [Test]
    public Task PrivateOnPartialClassNestedInNonPartial_Reported()
    {
        var outcome = Run(
            """
            public class Crm
            {
                public partial class Customer
                {
                    [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                    public List<Line> Lines { get; set; } = new();

                    static void ConfigureLines(WordTableBuilder<Line> builder)
                    {
                    }
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Crm.Customer Customer { get; set; } = new();
            }
            """);

        return AssertReported(outcome, "ConfigureLines", "Sample.Crm.Customer");
    }

    // A model nested in the type that declares the method sees its privates, as any nested type
    // does.
    [Test]
    public async Task PrivateOnTypeEnclosingTheModel_CalledDirectly()
    {
        var outcome = Run(
            """
            public partial class Customer
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }

                [ParchmentBindable]
                public partial class Invoice
                {
                    public Customer Customer { get; set; } = new();
                }
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ConfigureLines(");
        await Assert.That(outcome.Generated).DoesNotContain("ParchmentConfigure_");
    }

    // nameof accepts a method the declaring type inherits, and so does the call. Protected is out
    // of reach from an unrelated model but not from the declaring type's own partial part.
    [Test]
    public async Task ProtectedInheritedByPartialDeclaringType_Forwarded()
    {
        var outcome = Run(
            """
            public class Party
            {
                protected static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            public partial class Customer : Party
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("partial class Customer");
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ParchmentConfigure_Sample_Invoice_ConfigureLines(");
    }

    [Test]
    public async Task InternalInheritedByNonPartialDeclaringType_CalledDirectly()
    {
        var outcome = Run(
            """
            public class Party
            {
                internal static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            public class Customer : Party
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("b => global::Sample.Customer.ConfigureLines(");
    }

    // Overloads stay overloads: one forwarder each under one name, so the cast at the call site
    // chooses among them as it would among the methods.
    [Test]
    public async Task PrivateOverloads_EachForwarded()
    {
        var outcome = Run(
            """
            public class Note
            {
                public string Text { get; set; } = "";
            }

            public partial class Customer
            {
                [ExcelsiorTable(Configure = nameof(Configure))]
                public List<Line> Lines { get; set; } = new();

                [ExcelsiorTable(Configure = nameof(Configure))]
                public List<Note> Notes { get; set; } = new();

                static void Configure(WordTableBuilder<Line> builder)
                {
                }

                static void Configure(WordTableBuilder<Note> builder)
                {
                }
            }

            [ParchmentBindable]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """);

        await AssertCompiles(outcome);
        await Assert.That(outcome.Generated).Contains("ParchmentConfigure_Sample_Invoice_Configure(global::Excelsior.WordTableBuilder<global::Sample.Line> builder)");
        await Assert.That(outcome.Generated).Contains("ParchmentConfigure_Sample_Invoice_Configure(global::Excelsior.WordTableBuilder<global::Sample.Note> builder)");
    }

    // A type from another assembly cannot be given a part, and its private method cannot be made
    // reachable from here at all.
    [Test]
    public Task PrivateInAnotherAssembly_Reported()
    {
        var library = CSharpCompilation.Create(
            "Library",
            [
                CSharpSyntaxTree.ParseText(
                    """
                    namespace Library;

                    public class DocumentBase
                    {
                        static void ConfigureLines(object builder)
                        {
                        }
                    }
                    """)
            ],
            [MetadataReference.CreateFromFile(typeof(object).Assembly.Location)],
            new(OutputKind.DynamicallyLinkedLibrary));

        var outcome = Run(
            """
            [ParchmentBindable]
            public partial class Invoice : Library.DocumentBase
            {
                [ExcelsiorTable(Configure = "ConfigureLines")]
                public List<Line> Lines { get; set; } = new();
            }
            """,
            library.ToMetadataReference());

        return AssertReported(outcome, "ConfigureLines", "Sample.Invoice");
    }

    [Test]
    public async Task DocxModel_ForwarderGoesInItsOwnFile()
    {
        var source =
            prelude +
            """
            public partial class Customer
            {
                [ExcelsiorTable(Configure = nameof(ConfigureLines))]
                public List<Line> Lines { get; set; } = new();

                static void ConfigureLines(WordTableBuilder<Line> builder)
                {
                }
            }

            [ParchmentModel]
            public partial class Invoice
            {
                public Customer Customer { get; set; } = new();
            }
            """;
        var result = GeneratorDriver.Run(source, "{{ Customer.Lines }}").Results.Single();

        await Assert.That(result.Diagnostics).IsEmpty();
        var hints = result.GeneratedSources.Select(_ => _.HintName).ToList();
        await Assert.That(hints).Contains("Sample_Invoice_ParchmentModel.g.cs");
        await Assert.That(hints).Contains("Sample_Invoice_ParchmentModel_Configure.g.cs");
    }
}
