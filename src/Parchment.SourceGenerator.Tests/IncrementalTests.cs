public class IncrementalTests
{
    const string source =
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

    [Test]
    public async Task PipelineCachesWhenUnrelatedSyntaxAdded()
    {
        var setup = GeneratorDriver.CreateDriver(source, "Hello {{ Customer.Name }}!");
        var driver = (CSharpGeneratorDriver) setup.Driver.RunGenerators(setup.Compilation);

        var compilation2 = setup.Compilation.AddSyntaxTrees(
            CSharpSyntaxTree.ParseText("namespace Unrelated; public class Noise { public int X { get; set; } }"));
        driver = (CSharpGeneratorDriver) driver.RunGenerators(compilation2);

        var runResult = driver.GetRunResult().Results.Single();

        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.TargetsCollected);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.DocsCollected);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.Combined);
        await AssertOutputsNotModified(runResult);
    }

    // The model's syntax tree rides along with its snapshot so diagnostics land at a source
    // location. The tree only changes identity when its own file is edited, so an edit to another
    // file must leave the pipeline cached and the reported diagnostic in source.
    [Test]
    public async Task PipelineCachesWhenAnotherFileEdited()
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Letter.parchment.docx", GeneratorDriver.BuildDocxBytesWithoutPrivacyFlag("Hello {{ Customer.Name }}!")));
        var compilation1 = setup.Compilation.AddSyntaxTrees(
            CSharpSyntaxTree.ParseText("namespace Unrelated; public class Noise { public int X { get; set; } }"));
        var driver = (CSharpGeneratorDriver) setup.Driver.RunGenerators(compilation1);

        var compilation2 = compilation1.ReplaceSyntaxTree(
            compilation1.SyntaxTrees[2],
            CSharpSyntaxTree.ParseText("namespace Unrelated; public class Noise { public int X { get; set; } public int Y { get; set; } }"));
        driver = (CSharpGeneratorDriver) driver.RunGenerators(compilation2);

        var runResult = driver.GetRunResult().Results.Single();

        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.Targets);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.TargetsCollected);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.Combined);
        await AssertOutputsNotModified(runResult);

        var diagnostic = runResult.Diagnostics.Single();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH012");
        await Assert.That(diagnostic.Location.SourceTree).IsSameReferenceAs(compilation2.SyntaxTrees[1]);
    }

    // Three models reaching one type declared in a file of its own. Its entry is built once per
    // compilation and handed to all three (ShapeBuilder.sharedEntries), which must cost neither
    // guarantee: an edit elsewhere leaves every model cached, and an edit to the shared type
    // reaches every model.
    const string sharingModels =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Letter
        {
            public Customer Customer { get; set; } = new();
        }

        [ParchmentBindable]
        public partial class Invoice
        {
            public Customer Customer { get; set; } = new();
        }

        [ParchmentBindable]
        public partial class Receipt
        {
            public Customer Customer { get; set; } = new();
        }
        """;

    static (CSharpGeneratorDriver driver, CSharpCompilation compilation) RunSharingModels()
    {
        var setup = GeneratorDriver.CreateDriver(sharingModels, "Hello {{ Customer.Name }}!");
        var compilation = setup.Compilation.AddSyntaxTrees(
            CSharpSyntaxTree.ParseText("""namespace Sample; public class Customer { public string Name { get; set; } = ""; }"""),
            CSharpSyntaxTree.ParseText("namespace Unrelated; public class Noise { public int X { get; set; } }"));
        return ((CSharpGeneratorDriver) setup.Driver.RunGenerators(compilation), compilation);
    }

    [Test]
    public async Task SharedTypeLeavesEveryModelCachedWhenAnotherFileEdited()
    {
        var (driver, compilation) = RunSharingModels();

        var compilation2 = compilation.ReplaceSyntaxTree(
            compilation.SyntaxTrees[3],
            CSharpSyntaxTree.ParseText("namespace Unrelated; public class Noise { public int X { get; set; } public int Y { get; set; } }"));
        driver = (CSharpGeneratorDriver) driver.RunGenerators(compilation2);

        var runResult = driver.GetRunResult().Results.Single();

        await Assert.That(runResult.GeneratedSources.Length).IsEqualTo(3);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.Targets);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.Bindables);
        await AssertNotModified(runResult, ParchmentTemplateGenerator.Stages.Combined);
        await AssertOutputsNotModified(runResult);
    }

    [Test]
    public async Task SharedTypeEditRegeneratesEveryModelReachingIt()
    {
        var (driver, compilation) = RunSharingModels();
        foreach (var generated in driver.GetRunResult().Results.Single().GeneratedSources)
        {
            await Assert.That(generated.SourceText.ToString()).DoesNotContain("Email");
        }

        var compilation2 = compilation.ReplaceSyntaxTree(
            compilation.SyntaxTrees[2],
            CSharpSyntaxTree.ParseText("""namespace Sample; public class Customer { public string Name { get; set; } = ""; public string Email { get; set; } = ""; }"""));
        driver = (CSharpGeneratorDriver) driver.RunGenerators(compilation2);

        var runResult = driver.GetRunResult().Results.Single();

        await Assert.That(runResult.GeneratedSources.Length).IsEqualTo(3);
        foreach (var generated in runResult.GeneratedSources)
        {
            await Assert.That(generated.SourceText.ToString()).Contains("\"Email\"");
        }

        await AssertAllModified(runResult, ParchmentTemplateGenerator.Stages.Targets);
        await AssertAllModified(runResult, ParchmentTemplateGenerator.Stages.Bindables);
    }

    static async Task AssertAllModified(GeneratorRunResult runResult, string stageName)
    {
        foreach (var step in runResult.TrackedSteps[stageName])
        {
            foreach (var output in step.Outputs)
            {
                await Assert.That(output.Reason).IsEqualTo(IncrementalStepRunReason.Modified);
            }
        }
    }

    [Test]
    public Task PipelineReRunsWhenAttributedClassChanges()
    {
        var setup = GeneratorDriver.CreateDriver(source, "Hello {{ Customer.Name }}!");
        var driver = (CSharpGeneratorDriver) setup.Driver.RunGenerators(setup.Compilation);

        // Edit the attributed class's source so the extract stage must re-run and produce a
        // new TargetInfo — this proves the "Cached" assertion above isn't just a false positive
        // from the pipeline never running at all.
        var editedSource = source.Replace("class Letter", "class LetterV2");
        var compilation2 = setup.Compilation.ReplaceSyntaxTree(
            setup.Compilation.SyntaxTrees[1],
            CSharpSyntaxTree.ParseText(editedSource));
        driver = (CSharpGeneratorDriver) driver.RunGenerators(compilation2);

        var runResult = driver.GetRunResult().Results.Single();
        return AssertAnyModified(runResult, ParchmentTemplateGenerator.Stages.TargetsCollected);
    }

    [Test]
    public async Task PipelineReRunsWhenDocxParagraphsChange()
    {
        var setup = GeneratorDriver.CreateDriver(source, "Hello {{ Customer.Name }}!");
        var driver = (CSharpGeneratorDriver) setup.Driver.RunGenerators(setup.Compilation);

        var rewritten = GeneratorDriver.RewriteDocx(setup.DocxPath, "Goodbye {{ Customer.Name }}!");
        driver = (CSharpGeneratorDriver) driver.ReplaceAdditionalText(setup.DocxAdditionalText, rewritten);
        driver = (CSharpGeneratorDriver) driver.RunGenerators(setup.Compilation);

        var runResult = driver.GetRunResult().Results.Single();
        await AssertAnyModified(runResult, ParchmentTemplateGenerator.Stages.DocsCollected);
        await AssertAnyModified(runResult, ParchmentTemplateGenerator.Stages.Combined);
    }

    // The template's bytes are embedded into the generated source, so a resave that changes bytes
    // without changing a paragraph (rsid churn, package metadata) must re-emit — the embedded
    // copy would otherwise go stale. This is a deliberate reversal of the old paragraph-level
    // caching, which was valid only while the runtime read the template from disk.
    [Test]
    public async Task PipelineReRunsWhenDocxRewrittenWithSameParagraphs()
    {
        var setup = GeneratorDriver.CreateDriver(source, "Hello {{ Customer.Name }}!");
        var driver = (CSharpGeneratorDriver) setup.Driver.RunGenerators(setup.Compilation);

        var rewritten = GeneratorDriver.RewriteDocx(setup.DocxPath, "Hello {{ Customer.Name }}!");
        driver = (CSharpGeneratorDriver) driver.ReplaceAdditionalText(setup.DocxAdditionalText, rewritten);
        driver = (CSharpGeneratorDriver) driver.RunGenerators(setup.Compilation);

        // OPC packages are not byte-reproducible (relationship ids draw fresh randomness), so the
        // rewrite always lands new bytes and the docs stage must notice.
        var runResult = driver.GetRunResult().Results.Single();
        await AssertAnyModified(runResult, ParchmentTemplateGenerator.Stages.DocsCollected);
        await AssertAnyModified(runResult, ParchmentTemplateGenerator.Stages.Combined);
    }

    static async Task AssertNotModified(GeneratorRunResult runResult, string stageName)
    {
        var steps = runResult.TrackedSteps[stageName];
        foreach (var step in steps)
        {
            foreach (var output in step.Outputs)
            {
                await Assert.That(output.Reason)
                    .IsEqualTo(IncrementalStepRunReason.Cached)
                    .Or.IsEqualTo(IncrementalStepRunReason.Unchanged);
            }
        }
    }

    static async Task AssertAnyModified(GeneratorRunResult runResult, string stageName)
    {
        var steps = runResult.TrackedSteps[stageName];
        var anyModified = steps
            .SelectMany(_ => _.Outputs)
            .Any(_ => _.Reason is IncrementalStepRunReason.Modified or IncrementalStepRunReason.New);
        await Assert.That(anyModified).IsTrue();
    }

    static async Task AssertOutputsNotModified(GeneratorRunResult runResult)
    {
        foreach (var pair in runResult.TrackedOutputSteps)
        {
            foreach (var step in pair.Value)
            {
                foreach (var output in step.Outputs)
                {
                    await Assert.That(output.Reason)
                        .IsEqualTo(IncrementalStepRunReason.Cached)
                        .Or.IsEqualTo(IncrementalStepRunReason.Unchanged);
                }
            }
        }
    }
}