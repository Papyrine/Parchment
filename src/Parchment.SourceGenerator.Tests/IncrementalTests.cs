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

    // Note: we intentionally do not test "editing the model class in a separate file
    // re-runs validation" — ForAttributeWithMetadataName only re-fires on syntax changes to
    // the attributed class itself, so this scenario is a known gap documented in
    // ParchmentTemplateGenerator.Initialize and ShapeBuilder.
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