using System.Reflection;
using System.Runtime.Loader;

// What the generator does with each way a model can be declared. The snapshots elsewhere show what
// is emitted; these compile it, because the failure being guarded against is generated source that
// the generator was happy with and the compiler is not. The supported shapes are then loaded and
// their module initializer run, since an initializer that compiles but is never reached would
// leave the model unregistered with nothing to say so.
public class ModelDeclarationGeneratorTests
{
    // Stand-ins for the parts of Parchment and Fluid the generated source calls — this project
    // references neither. Each registration records what it was handed.
    const string runtimeStubs =
        """
        namespace Fluid
        {
            public interface IMemberAccessor;
        }

        namespace Fluid.Accessors
        {
            public sealed class DelegateAccessor(System.Func<object, string, object?> getter) :
                Fluid.IMemberAccessor
            {
                public System.Func<object, string, object?> Getter { get; } = getter;
            }
        }

        namespace Parchment.Generated
        {
            public static class GeneratedRegistration
            {
                public static readonly System.Collections.Generic.List<string> Registered = [];

                public static void RegisterDocxTemplate(System.Type modelType, byte[] template, Parchment.ProtectionMode protection = Parchment.ProtectionMode.WhenEditable) =>
                    Registered.Add($"docx:{modelType.FullName}");

                public static void RegisterMarkdownTemplate(System.Type modelType, string markdown, byte[]? styleSource = null) =>
                    Registered.Add($"markdown:{modelType.FullName}");

                public static void RegisterFluidAccessors(
                    System.Type type,
                    System.Collections.Generic.IEnumerable<System.Collections.Generic.KeyValuePair<string, Fluid.IMemberAccessor>> accessors) =>
                    Registered.Add($"accessors:{type.FullName}");
            }
        }
        """;

    sealed record Outcome(
        ImmutableArray<Diagnostic> Generator,
        ImmutableArray<Diagnostic> Compile,
        int GeneratedTrees,
        List<string> Registered);

    static Outcome Run(string source, params TemplateFile[] templates)
    {
        if (templates.Length == 0)
        {
            templates = [Docx("M")];
        }

        // Given a path because a file-local type takes part of its metadata name from it.
        var setup = GeneratorDriver.CreateDriverWithDocxes(source, templates, modelFileName: "Model.cs");
        var compilation = setup.Compilation.AddSyntaxTrees(
            CSharpSyntaxTree.ParseText(runtimeStubs, (CSharpParseOptions)setup.Compilation.SyntaxTrees[0].Options));
        var driver = setup.Driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out _);
        var result = driver.GetRunResult();

        var compile = output.GetDiagnostics()
            .Where(_ => _.Severity >= DiagnosticSeverity.Warning)
            .ToImmutableArray();

        var registered = new List<string>();
        if (compile.IsEmpty)
        {
            using var stream = new MemoryStream();
            var emitted = output.Emit(stream);
            if (!emitted.Success)
            {
                throw new InvalidOperationException(string.Join(Environment.NewLine, emitted.Diagnostics));
            }

            stream.Position = 0;
            var context = new AssemblyLoadContext(null, isCollectible: true);
            try
            {
                var assembly = context.LoadFromStream(stream);
                RuntimeHelpers.RunModuleConstructor(assembly.ManifestModule.ModuleHandle);
                var field = assembly
                    .GetType("Parchment.Generated.GeneratedRegistration", throwOnError: true)!
                    .GetField("Registered", BindingFlags.Public | BindingFlags.Static)!;
                registered.AddRange((List<string>)field.GetValue(null)!);
            }
            finally
            {
                context.Unload();
            }
        }

        return new(result.Results.Single().Diagnostics, compile, result.GeneratedTrees.Length, registered);
    }

    static TemplateFile Docx(string typeName) =>
        new($"{typeName}.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }}"));

    static async Task AssertRegisters(string source, params string[] expected)
    {
        var outcome = Run(source);
        await Assert.That(Describe(outcome.Generator)).IsEqualTo("");
        await Assert.That(Describe(outcome.Compile)).IsEqualTo("");
        foreach (var registration in expected)
        {
            await Assert.That(outcome.Registered).Contains(registration);
        }
    }

    static async Task AssertUnsupported(string source, string modelDisplayName, string reason)
    {
        var outcome = Run(source);
        await Assert.That(outcome.Generator.Length).IsEqualTo(1);
        var diagnostic = outcome.Generator[0];
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH031");
        await Assert.That(diagnostic.GetMessage()).Contains($"'{modelDisplayName}'");
        await Assert.That(diagnostic.GetMessage()).Contains(reason);
        // The point of the diagnostic: one error at the model, none from generated source.
        await Assert.That(outcome.GeneratedTrees).IsEqualTo(0);
        await Assert.That(Describe(outcome.Compile)).IsEqualTo("");
    }

    static string Describe(ImmutableArray<Diagnostic> diagnostics) =>
        string.Join(Environment.NewLine, diagnostics.Select(_ => $"{_.Id}: {_.GetMessage()}"));

    [Test]
    [Arguments("private")]
    [Arguments("protected")]
    [Arguments("private protected")]
    [Arguments("protected internal")]
    [Arguments("internal")]
    public Task NestedModel_AnyAccessibility_Registers(string accessibility) =>
        AssertRegisters(
            $$"""
              using Parchment;

              namespace Sample;

              public partial class Outer
              {
                  [ParchmentModel]
                  {{accessibility}} partial class M
                  {
                      public string Name { get; set; } = "";
                  }
              }
              """,
            "accessors:Sample.Outer+M",
            "docx:Sample.Outer+M");

    [Test]
    public Task PrivateNestedBindable_Registers() =>
        AssertRegisters(
            """
            using Parchment;

            namespace Sample;

            public partial class Outer
            {
                [ParchmentBindable]
                private partial class M
                {
                    public string Name { get; set; } = "";
                }
            }
            """,
            "accessors:Sample.Outer+M");

    [Test]
    public async Task PrivateNestedModel_MarkdownTemplate_Registers()
    {
        var outcome = Run(
            """
            using Parchment;

            namespace Sample;

            public partial class Outer
            {
                [ParchmentModel]
                private partial class M
                {
                    public string Name { get; set; } = "";
                }
            }
            """,
            new TemplateFile("M.parchment.md", "Hello {{ Name }}"u8.ToArray()));
        await Assert.That(Describe(outcome.Generator)).IsEqualTo("");
        await Assert.That(Describe(outcome.Compile)).IsEqualTo("");
        await Assert.That(outcome.Registered).Contains("markdown:Sample.Outer+M");
    }

    [Test]
    public Task PrivateModel_ReachingAPrivateSiblingType_Registers() =>
        AssertRegisters(
            """
            using Parchment;

            namespace Sample;

            public partial class Outer
            {
                private class Customer
                {
                    public string Name { get; set; } = "";
                }

                [ParchmentModel]
                private partial class M
                {
                    public string Name { get; set; } = "";
                    public Customer Customer { get; set; } = new();
                }
            }
            """,
            "accessors:Sample.Outer+Customer",
            "docx:Sample.Outer+M");

    // One forwarder per private level: a private type can only be named from the type declaring it.
    [Test]
    [Arguments("private", "private")]
    [Arguments("private", "public")]
    [Arguments("public", "private")]
    [Arguments("protected", "private protected")]
    public Task DoublyNestedModel_Registers(string middle, string model) =>
        AssertRegisters(
            $$"""
              using Parchment;

              namespace Sample;

              public partial class Outer
              {
                  {{middle}} partial class Middle
                  {
                      [ParchmentModel]
                      {{model}} partial class M
                      {
                          public string Name { get; set; } = "";
                      }
                  }
              }
              """,
            "docx:Sample.Outer+Middle+M");

    [Test]
    public Task PrivateLevelsEitherSideOfAPublicOne_Registers() =>
        AssertRegisters(
            """
            using Parchment;

            namespace Sample;

            public partial class Outer
            {
                private partial class First
                {
                    public partial class Second
                    {
                        [ParchmentModel]
                        private partial class M
                        {
                            public string Name { get; set; } = "";
                        }
                    }
                }
            }
            """,
            "docx:Sample.Outer+First+Second+M");

    [Test]
    public Task TwoPrivateModelsWithOneNameUnderOneType_Register() =>
        AssertRegisters(
            """
            using Parchment;

            namespace Sample;

            public partial class Outer
            {
                private partial class First
                {
                    [ParchmentModel]
                    private partial class M
                    {
                        public string Name { get; set; } = "";
                    }
                }

                private partial class Second
                {
                    [ParchmentModel]
                    private partial class M
                    {
                        public string Name { get; set; } = "";
                    }
                }

                [ParchmentModel]
                private partial class M
                {
                    public string Name { get; set; } = "";
                }
            }
            """,
            "docx:Sample.Outer+First+M",
            "docx:Sample.Outer+Second+M",
            "docx:Sample.Outer+M");

    [Test]
    [Arguments("partial interface")]
    [Arguments("partial struct")]
    [Arguments("partial record struct")]
    [Arguments("partial record")]
    [Arguments("static partial class")]
    [Arguments("abstract partial class")]
    public async Task ModelNestedInAnyKindOfType_Registers(string enclosing)
    {
        foreach (var accessibility in (string[])["public", "private"])
        {
            await AssertRegisters(
                $$"""
                  using Parchment;

                  namespace Sample;

                  public {{enclosing}} Outer
                  {
                      [ParchmentModel]
                      {{accessibility}} partial class M
                      {
                          public string Name { get; set; } = "";
                      }
                  }
                  """,
                "docx:Sample.Outer+M");
        }
    }

    [Test]
    [Arguments("abstract partial class")]
    [Arguments("sealed partial class")]
    [Arguments("static partial class")]
    [Arguments("partial record")]
    public Task ModelModifiers_Register(string declaration) =>
        AssertRegisters(
            $$"""
              using Parchment;

              namespace Sample;

              [ParchmentModel]
              public {{declaration}} M
              {
                  public static string Name { get; set; } = "";
              }
              """,
            "docx:Sample.M");

    [Test]
    public Task GlobalNamespace_PrivateNestedModel_Registers() =>
        AssertRegisters(
            """
            using Parchment;

            public partial class Outer
            {
                [ParchmentModel]
                private partial class M
                {
                    public string Name { get; set; } = "";
                }
            }
            """,
            "docx:Outer+M");

    [Test]
    [Arguments("ParchmentModel")]
    [Arguments("ParchmentBindable")]
    public Task GenericModel_ReportsParch031(string attribute) =>
        AssertUnsupported(
            $$"""
              using Parchment;

              namespace Sample;

              [{{attribute}}]
              public partial class M<T>
              {
                  public string Name { get; set; } = "";
                  public T? Value { get; set; }
              }
              """,
            "M",
            "it is generic");

    [Test]
    [Arguments("public")]
    [Arguments("private")]
    public Task ModelNestedInGenericType_ReportsParch031(string accessibility) =>
        AssertUnsupported(
            $$"""
              using Parchment;

              namespace Sample;

              public partial class Outer<T>
              {
                  [ParchmentModel]
                  {{accessibility}} partial class M
                  {
                      public string Name { get; set; } = "";
                  }
              }
              """,
            "Outer.M",
            "nested inside the generic type 'Outer'");

    [Test]
    [Arguments("ParchmentModel")]
    [Arguments("ParchmentBindable")]
    public Task FileLocalModel_ReportsParch031(string attribute) =>
        AssertUnsupported(
            $$"""
              using Parchment;

              namespace Sample;

              [{{attribute}}]
              file partial class M
              {
                  public string Name { get; set; } = "";
              }
              """,
            "M",
            "it is file-local");

    [Test]
    public Task ModelNestedInFileLocalType_ReportsParch031() =>
        AssertUnsupported(
            """
            using Parchment;

            namespace Sample;

            file partial class Outer
            {
                [ParchmentModel]
                public partial class M
                {
                    public string Name { get; set; } = "";
                }
            }
            """,
            "Outer.M",
            "nested inside the file-local type 'Outer'");

    // Nothing a partial modifier on the enclosing type would fix, so the shape is what is reported.
    [Test]
    public async Task UnsupportedShapeInsideNonPartialType_ReportsOnlyParch031()
    {
        var outcome = Run(
            """
            using Parchment;

            namespace Sample;

            public class Outer<T>
            {
                [ParchmentModel]
                public partial class M
                {
                    public string Name { get; set; } = "";
                }
            }
            """);
        await Assert.That(outcome.Generator.Select(_ => _.Id)).IsEquivalentTo((string[])["PARCH031"]);
        await Assert.That(outcome.GeneratedTrees).IsEqualTo(0);
    }
}
