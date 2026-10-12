/// <summary>
/// Runs the generator and then emits what it produced, with the model, against the real
/// Parchment assembly. <see cref="GeneratorDriver.Compile"/> binds the generated source but never
/// emits it, so output the compiler cannot write to an assembly passes it unnoticed, and there is
/// no assembly to load and render from.
/// </summary>
static class CompiledGenerator
{
    static readonly Lazy<MetadataReference[]> references = new(BuildReferences);

    // The trusted platform assemblies are the framework plus everything the test host deployed,
    // which — through the project reference — includes Parchment and what it depends on.
    static MetadataReference[] BuildReferences()
    {
        var tpa = (string?) AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? "";
        return tpa
            .Split(Path.PathSeparator)
            .Where(_ => !string.IsNullOrEmpty(_))
            .Select(MetadataReference (_) => MetadataReference.CreateFromFile(_))
            .ToArray();
    }

    public sealed record Result(ImmutableArray<Diagnostic> Errors, byte[] Assembly, string Generated);

    public static Result Run(string source, TemplateFile[] files, LanguageVersion language = LanguageVersion.Default)
    {
        var directory = Path.Combine(Path.GetTempPath(), "parchment-sg-tests", Guid.NewGuid().ToString("N"));
        var texts = ImmutableArray.CreateBuilder<AdditionalText>();
        foreach (var (name, bytes) in files)
        {
            var path = Path.Combine(directory, name);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, bytes);
            texts.Add(new GeneratorDriver.PathAdditionalText(path));
        }

        var parse = new CSharpParseOptions(language);
        var compilation = CSharpCompilation.Create(
            $"CompiledGenerator{Guid.NewGuid():N}",
            [CSharpSyntaxTree.ParseText(source, parse)],
            references.Value,
            new(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));

        var driver = CSharpGeneratorDriver.Create(
            generators: [new ParchmentTemplateGenerator().AsSourceGenerator()],
            additionalTexts: texts.ToImmutable(),
            parseOptions: parse);

        var ran = driver.RunGeneratorsAndUpdateCompilation(compilation, out var output, out var generatorDiagnostics);

        using var stream = new MemoryStream();
        var emitted = output.Emit(stream);

        var errors = generatorDiagnostics
            .Concat(ran.GetRunResult().Diagnostics)
            .Concat(emitted.Diagnostics)
            .Where(_ => _.Severity == DiagnosticSeverity.Error)
            .Distinct()
            .ToImmutableArray();
        var generated = string.Concat(ran.GetRunResult().GeneratedTrees.Select(_ => _.GetText().ToString()));
        return new(errors, stream.ToArray(), generated);
    }
}
