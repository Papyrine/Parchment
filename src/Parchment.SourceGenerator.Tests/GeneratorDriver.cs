static class GeneratorDriver
{
    const string attributeSource =
        """
        namespace Parchment
        {
            [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
            public sealed class ParchmentModelAttribute : System.Attribute
            {
                public ProtectionMode Protection { get; set; }
            }

            [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
            public sealed class ParchmentBindableAttribute : System.Attribute
            {
            }

            public enum ProtectionMode
            {
                WhenEditable,
                None
            }

            public sealed class TemplateStore { }

            [System.AttributeUsage(System.AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
            public sealed class ExcelsiorTableAttribute : System.Attribute { }

            [System.AttributeUsage(System.AttributeTargets.Property | System.AttributeTargets.Field, AllowMultiple = false, Inherited = false)]
            public sealed class EditableFieldAttribute : System.Attribute
            {
                public bool MultiLine { get; set; }

                public string? DateFormat { get; set; }
            }
        }
        """;

    public static GeneratorDriverRunResult Run(string userSource, params string[] templateParagraphs)
    {
        var setup = CreateDriver(userSource, templateParagraphs);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }

    // The convention names the template after the model type, plus the .parchment marker, so the
    // default file name is read off the [ParchmentModel] target in the source under test.
    public static DriverSetup CreateDriver(string userSource, params string[] templateParagraphs) =>
        CreateDriverWithDocxes(userSource, new TemplateFile($"{ModelTypeName(userSource)}.parchment.docx", BuildDocx(templateParagraphs)));

    /// <summary>
    /// The name of the first <c>[ParchmentModel]</c>-decorated type in the source under test —
    /// what the convention will look for a template file to match.
    /// </summary>
    public static string ModelTypeName(string userSource)
    {
        var match = Regex.Match(
            userSource,
            @"\[ParchmentModel[^\]]*\]\s*(?:\[[^\]]*\]\s*)*(?:public\s+|internal\s+)*partial\s+(?:record\s+struct|class|record|struct)\s+(\w+)");
        if (!match.Success)
        {
            throw new InvalidOperationException("No [ParchmentModel] partial type found in the source under test.");
        }

        return match.Groups[1].Value;
    }

    public static DriverSetup CreateDriverWithDocxes(
        string userSource,
        params TemplateFile[] docxes) =>
        CreateDriverWithDocxes(userSource, docxes, modelFileName: null);

    /// <summary>
    /// Runs the generator as a real build does — the files under a project directory the compiler
    /// can see — so diagnostic messages carry project-relative paths.
    /// </summary>
    public static GeneratorDriverRunResult RunInProject(
        string userSource,
        params TemplateFile[] files)
    {
        var setup = CreateDriverWithDocxes(
            userSource,
            files,
            modelFileName: "Model.cs",
            withProjectDir: true);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }

    public static DriverSetup CreateDriverWithDocxes(
        string userSource,
        TemplateFile[] docxes,
        string? modelFileName,
        bool withProjectDir = false)
    {
        var directory = Path.Combine(Path.GetTempPath(), "parchment-sg-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);

        var texts = ImmutableArray.CreateBuilder<AdditionalText>();
        var paths = ImmutableArray.CreateBuilder<string>();
        foreach (var (name, bytes) in docxes)
        {
            var path = Path.Combine(directory, name);
            // A template may be given a nested path, so its folder has to exist first.
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, bytes);
            texts.Add(new PathAdditionalText(path));
            paths.Add(path);
        }

        var syntaxTrees = new[]
        {
            CSharpSyntaxTree.ParseText(attributeSource),
            modelFileName == null
                ? CSharpSyntaxTree.ParseText(userSource)
                : CSharpSyntaxTree.ParseText(userSource, path: CreateModelPath(directory, modelFileName))
        };

        var compilation = CSharpCompilation.Create(
            "GeneratorTest",
            syntaxTrees,
            BuildReferences(),
            // Nullable context on so `string?` members carry NullableAnnotation.Annotated —
            // ShapeBuilder reads it to decide editable-field nullability.
            new(OutputKind.DynamicallyLinkedLibrary, nullableContextOptions: NullableContextOptions.Enable));

        var additionalTexts = texts.ToImmutable();
        var driver = CSharpGeneratorDriver.Create(
            generators: [new ParchmentTemplateGenerator().AsSourceGenerator()],
            additionalTexts: additionalTexts,
            parseOptions: (CSharpParseOptions) syntaxTrees[0].Options,
            optionsProvider: withProjectDir ? new TestOptionsProvider(directory + Path.DirectorySeparatorChar) : null,
            driverOptions: new(IncrementalGeneratorOutputKind.None, trackIncrementalGeneratorSteps: true));

        return new(driver, compilation, additionalTexts, paths.ToImmutable());
    }

    static string CreateModelPath(string directory, string modelFileName)
    {
        var path = Path.Combine(directory, modelFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        return path;
    }

    public static byte[] BuildDocxBytes(params string[] paragraphs) => BuildDocx(paragraphs);

    /// <summary>
    /// A docx whose body and default header both carry paragraphs — for validating that
    /// body-scoped editable rules ignore header occurrences (the read-only mirror pattern).
    /// </summary>
    public static byte[] BuildDocxBytesWithHeader(string[] bodyParagraphs, string[] headerParagraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            var header = new Header();
            foreach (var text in headerParagraphs)
            {
                header.Append(BuildParagraph(text));
            }

            headerPart.Header = header;

            foreach (var text in bodyParagraphs)
            {
                body.Append(BuildParagraph(text));
            }

            body.Append(
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    }));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new(
                new RemovePersonalInformation(),
                new RemoveDateAndTime());
        }

        return stream.ToArray();
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    public static byte[] BuildDocxBytesWithoutPrivacyFlag(params string[] paragraphs) =>
        BuildDocx(paragraphs, removePersonalInformation: false);

    public static AdditionalText RewriteDocx(string path, params string[] paragraphs)
    {
        File.WriteAllBytes(path, BuildDocx(paragraphs));
        return new PathAdditionalText(path);
    }

    public static GeneratorDriverRunResult RunMarkdown(string userSource, string markdown, string? fileName = null)
    {
        var setup = CreateDriverWithFiles(
            userSource,
            new TemplateFile(fileName ?? $"{ModelTypeName(userSource)}.parchment.md", Encoding.UTF8.GetBytes(markdown)));
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }

    public static DriverSetup CreateDriverWithFiles(
        string userSource,
        params TemplateFile[] files) =>
        CreateDriverWithDocxes(userSource, files);

    public sealed record DriverSetup(
        CSharpGeneratorDriver Driver,
        CSharpCompilation Compilation,
        ImmutableArray<AdditionalText> AdditionalTexts,
        ImmutableArray<string> DocxPaths)
    {
        public AdditionalText DocxAdditionalText => AdditionalTexts[0];
        public string DocxPath => DocxPaths[0];
    }

    static string WriteDocx(string[] paragraphs)
    {
        var directory = Path.Combine(Path.GetTempPath(), "parchment-sg-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "template.docx");
        File.WriteAllBytes(path, BuildDocx(paragraphs));
        return path;
    }

    static byte[] BuildDocx(string[] paragraphs, bool removePersonalInformation = true)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;
            foreach (var text in paragraphs)
            {
                body.Append(
                    new Paragraph(
                        new Run(new Text(text)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        })));
            }

            if (removePersonalInformation)
            {
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new(
                    new RemovePersonalInformation(),
                    new RemoveDateAndTime());
            }
        }

        return stream.ToArray();
    }

    static MetadataReference[] BuildReferences()
    {
        var tpa = (string?) AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? "";
        return tpa
            .Split(Path.PathSeparator)
            .Where(_ => !string.IsNullOrEmpty(_))
            .Select(MetadataReference (_) => MetadataReference.CreateFromFile(_))
            .ToArray();
    }

    // Stands in for what the SDK writes into the generated editorconfig: the project directory as
    // a compiler-visible property.
    sealed class TestOptionsProvider(string? projectDir) :
        AnalyzerConfigOptionsProvider
    {
        public override AnalyzerConfigOptions GlobalOptions { get; } = new Options(projectDir);

        public override AnalyzerConfigOptions GetOptions(SyntaxTree tree) => new Options(null);

        public override AnalyzerConfigOptions GetOptions(AdditionalText textFile) => new Options(null);

        sealed class Options(string? projectDir) :
            AnalyzerConfigOptions
        {
            public override bool TryGetValue(string key, [NotNullWhen(true)] out string? value)
            {
                if (projectDir != null &&
                    key == "build_property.ProjectDir")
                {
                    value = projectDir;
                    return true;
                }

                value = null;
                return false;
            }
        }
    }

    sealed class PathAdditionalText(string path) :
        AdditionalText
    {
        public override string Path { get; } = path;

        public override SourceText? GetText(Cancel cancel = default)
        {
            // Mirror Roslyn: text-based AdditionalFiles (.md) are exposed as SourceText so the
            // SG can use the canonical GetText path. Binary files (.docx / .dotx) return null and
            // the SG must read them via Path with stream-based APIs.
            if (Path.EndsWith(".md", StringComparison.OrdinalIgnoreCase))
            {
                return SourceText.From(File.ReadAllText(Path));
            }

            return null;
        }
    }
}
