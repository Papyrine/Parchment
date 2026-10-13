/// <summary>
/// <c>[ParchmentBindable]</c>: accessors-only emission for models registered by hand against
/// templates the generator cannot see. There is no runtime reflection fallback, so this module
/// initializer is what makes a hand-registered model renderable at all.
/// </summary>
public class BindableGeneratorTests
{
    [Test]
    public Task EmitsAccessorsAndModuleInitializer()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            public class Customer
            {
                public string Name { get; set; } = "";
            }

            [ParchmentBindable]
            public partial class Letter
            {
                public Customer Customer { get; set; } = new();
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return Verify(result);
    }

    // The accessor tables are one lambda per member of every reachable type, and only the members a
    // template binds are ever invoked. Left in the model's own partial they would count as covered
    // code and measure the model graph rather than the consumer's tests — 3150 lines at 63% in one
    // repo, enough to drop it below its coverage gate. So they go in a nested type that carries the
    // exclusion, which the model's own partial could not without excluding the consumer's members.
    [Test]
    public async Task GeneratedMembersAreExcludedFromCoverage()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentBindable]
            public partial class Letter
            {
                public string Name { get; set; } = "";
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var emitted = string.Concat(
            setup.Driver.RunGenerators(setup.Compilation).GetRunResult()
                .Results.SelectMany(_ => _.GeneratedSources)
                .Select(_ => _.SourceText.ToString()));

        await Assert.That(emitted).Contains("[global::System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]");
        await Assert.That(emitted).Contains("internal static class ParchmentGenerated_Sample_Letter");
        // Inside the nested type rather than beside it.
        await Assert.That(emitted.IndexOf("internal static class ParchmentGenerated_Sample_Letter", StringComparison.Ordinal))
            .IsLessThan(emitted.IndexOf("_Accessors_", StringComparison.Ordinal));
    }

    // A class carrying both attributes gets one emission: [ParchmentModel] wins, since it already
    // registers the accessors alongside the embedded template.
    [Test]
    public async Task ParchmentModelWins()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            [ParchmentBindable]
            public partial class Letter
            {
                public string Name { get; set; } = "";
            }
            """;
        var result = GeneratorDriver.Run(source, "Hello {{ Name }}!");
        var hintNames = result.GeneratedTrees
            .Select(_ => Path.GetFileName(_.FilePath))
            .ToList();
        await Assert.That(hintNames).Contains("Sample_Letter_ParchmentModel.g.cs");
        await Assert.That(hintNames.Any(_ => _.Contains("ParchmentBindable"))).IsFalse();
    }

    // The editable shape rules fire for bindable targets too — the template kind is unknown at
    // compile time, and an invalid editable member is an authoring mistake regardless.
    [Test]
    public async Task ShapeRules_Fire()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentBindable]
            public partial class Order
            {
                [EditableField]
                public bool? Maybe { get; set; }
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var ids = result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
        await Assert.That(ids).Contains("PARCH013");
    }

    // A member type deriving from List<T> brings the list's indexer in through its base chain. An
    // accessor for it reads `((Bag)o).this[]`, which is not C#, and a template has no way to ask
    // for it anyway.
    [Test]
    public async Task ListDerivedMember_EmitsNoIndexerAccessor()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Line
            {
                public string Name { get; set; } = "";
            }

            public class Bag : List<Line>
            {
                public string Label { get; set; } = "";
            }

            [ParchmentBindable]
            public partial class Letter
            {
                public Bag Bag { get; set; } = new();
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var syntaxErrors = result.GeneratedTrees
            .SelectMany(_ => _.GetDiagnostics())
            .Select(_ => _.Id)
            .ToList();
        var generated = string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

        await Assert.That(syntaxErrors).IsEmpty();
        await Assert.That(generated).DoesNotContain("this[]");
        // What the list itself declares is not binding surface either: a template sees an array.
        await Assert.That(generated).DoesNotContain("\"Capacity\"");
        await Assert.That(generated).Contains("\"Label\"");
    }

    // An indexer the model declares itself is no more a member a template can name, and a property
    // whose getter is missing or private is one the generated code cannot read.
    [Test]
    public async Task DeclaredIndexer_EmitsNoAccessor()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentBindable]
            public partial class Letter
            {
                public string Name { get; set; } = "";

                public string this[int index] => Name;

                public string WriteOnly { set { } }

                public string Hidden { private get; set; } = "";
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var syntaxErrors = result.GeneratedTrees
            .SelectMany(_ => _.GetDiagnostics())
            .Select(_ => _.Id)
            .ToList();
        var generated = string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

        await Assert.That(syntaxErrors).IsEmpty();
        await Assert.That(generated).DoesNotContain("this[]");
        await Assert.That(generated).DoesNotContain("WriteOnly");
        await Assert.That(generated).DoesNotContain("Hidden");
        await Assert.That(generated).Contains("\"Name\"");
    }

    // Legal members no accessor can be written for: a ref struct or a pointer cannot become an
    // object (CS0029), and a static abstract interface member is reachable through a type parameter
    // only (CS8926). The compile itself is checked where the real Parchment is referenced, by the
    // Awkward model in Parchment.Tests' ReferenceValidatorTests.
    [Test]
    public async Task UnbindableMemberType_EmitsNoAccessor()
    {
        var source =
            """
            using System;
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public interface IThing
            {
                static abstract string Kind { get; }

                static virtual string Fallback => "";

                static string Default => "";

                string Label { get; }
            }

            public ref struct Cursor
            {
                public int Position;
            }

            [ParchmentBindable]
            public unsafe partial class Letter
            {
                int count;

                public string Name { get; set; } = "";

                public ref int Count => ref count;

                public ReadOnlySpan<char> Chars => Name.AsSpan();

                public Cursor Cursor => default;

                public int* Pointer;

                public delegate*<int, int> Function;

                public int*[] Pointers { get; set; } = [];

                public List<int*[]> PointerLists { get; set; } = [];

                public IThing? Thing { get; set; }
            }
            """;
        var generated = Generate(source);

        foreach (var name in new[] { "Chars", "Cursor", "Pointer", "Function", "Pointers", "PointerLists", "Kind", "Fallback" })
        {
            await Assert.That(generated).DoesNotContain($"\"{name}\"");
        }

        foreach (var name in new[] { "Name", "Count", "Thing", "Default", "Label" })
        {
            await Assert.That(generated).Contains($"\"{name}\"");
        }
    }

    // `dynamic` is not a legal operand of typeof (CS1962), and is `object` to the runtime, so that
    // is how every emitted type name spells it. A member or a tuple element that is merely *named*
    // dynamic is left alone.
    [Test]
    public async Task DynamicInAMemberType_IsWrittenAsObject()
    {
        var source =
            """
            using System.Collections.Generic;
            using Parchment;

            namespace Sample;

            public class Wrapper<T>
            {
                public T Value { get; set; } = default!;
            }

            [ParchmentBindable]
            public partial class Letter
            {
                public dynamic? Dyn { get; set; }

                public string dynamic { get; set; } = "";

                public Dictionary<string, dynamic> Map { get; set; } = new();

                public Wrapper<dynamic[]> Wrapped { get; set; } = new();

                public Wrapper<(int dynamic, dynamic B)> Pair { get; set; } = new();
            }
            """;
        var generated = Generate(source);

        await Assert.That(generated).DoesNotContain("<dynamic");
        await Assert.That(generated).DoesNotContain("dynamic>");
        await Assert.That(generated).DoesNotContain("typeof(dynamic)");
        await Assert.That(generated).Contains("typeof(global::System.Collections.Generic.KeyValuePair<string, object>)");
        await Assert.That(generated).Contains("typeof(global::Sample.Wrapper<object[]>)");
        await Assert.That(generated).Contains("typeof(global::Sample.Wrapper<(int dynamic, object B)>)");
        await Assert.That(generated).Contains("""new("dynamic", """);
    }

    // A name a derived type reuses for something that is not a readable property is what
    // `((Derived)o).Name` then binds to - a method group (CS8974), an event, a type - so the
    // inherited property is not reachable through the derived type. One hidden by a member the
    // generated code cannot see (protected) still is.
    [Test]
    public async Task MemberHiddenInADerivedType_EmitsNoAccessor()
    {
        var source =
            """
            using System;
            using Parchment;

            namespace Sample;

            public class Base
            {
                public string Kept { get; set; } = "";
                public string ByMethod { get; set; } = "";
                public string ByEvent { get; set; } = "";
                public string ByType { get; set; } = "";
                public string ByWriteOnly { get; set; } = "";
                public string BySpan { get; set; } = "";
                public string ByProtected { get; set; } = "";
                public string bycase { get; set; } = "";
            }

            public class Mid : Base
            {
                public new void ByMethod() { }
                public new event Action ByEvent { add { } remove { } }
                public new class ByType { }
                public new string ByWriteOnly { set { } }
                public new ReadOnlySpan<char> BySpan => default;
                protected new void ByProtected() { }
                public void ByCase() { }
            }

            [ParchmentBindable]
            public partial class Letter
            {
                public Mid Mid { get; set; } = new();
            }
            """;
        var generated = Generate(source);

        foreach (var name in new[] { "ByMethod", "ByEvent", "ByType", "ByWriteOnly", "BySpan" })
        {
            await Assert.That(generated).DoesNotContain($"\"{name}\"");
        }

        foreach (var name in new[] { "Kept", "ByProtected", "bycase" })
        {
            await Assert.That(generated).Contains($"\"{name}\"");
        }
    }

    static string Generate(string source)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));
    }

    // A base and a derived model can both be bindable — the type name is folded into the module
    // initializer's name so the derived one does not hide the base one (CS0108).
    [Test]
    public async Task InheritanceChain_DistinctInitializerNames()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentBindable]
            public partial class DocumentBase
            {
                public string Title { get; set; } = "";
            }

            [ParchmentBindable]
            public partial class Report : DocumentBase
            {
                public string Body { get; set; } = "";
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(source);
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        var generated = string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));
        await Assert.That(generated).Contains("InitializeParchmentBindableDocumentBase");
        await Assert.That(generated).Contains("InitializeParchmentBindableReport");
    }
}
