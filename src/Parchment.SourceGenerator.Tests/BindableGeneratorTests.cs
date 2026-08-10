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
        await Assert.That(emitted).Contains("internal static class ParchmentGeneratedLetter");
        // Inside the nested type rather than beside it.
        await Assert.That(emitted.IndexOf("internal static class ParchmentGeneratedLetter", System.StringComparison.Ordinal))
            .IsLessThan(emitted.IndexOf("_Accessors_", System.StringComparison.Ordinal));
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
