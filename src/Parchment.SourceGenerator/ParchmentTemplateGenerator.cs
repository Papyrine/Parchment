[Generator]
public sealed class ParchmentTemplateGenerator :
    IIncrementalGenerator
{
    const string attributeFullName = "Parchment.ParchmentModelAttribute";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // ForAttributeWithMetadataName only re-fires ExtractTarget when the attributed class's
        // own syntax changes. Editing a model class (e.g. adding/removing a property on Invoice)
        // in a separate file will NOT re-validate templates that reference it until the
        // attributed class is touched. The tradeoff: combining with CompilationProvider would
        // make the extract re-run every compilation and defeat the point of the primitive-only
        // pipeline below. Kicking the attributed file forces revalidation in the meantime.
        var targets = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                attributeFullName,
                static (node, _) => node is ClassDeclarationSyntax or RecordDeclarationSyntax,
                ExtractTarget)
            .Where(static _ => _ != null)
            .Select(static (target, _) => target!)
            .WithTrackingName(Stages.Targets)
            .Collect()
            .Select(static (array, _) => new EquatableArray<TargetInfo>(array))
            .WithTrackingName(Stages.TargetsCollected);

        var docs = context.AdditionalTextsProvider
            .Where(static _ => _.Path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            .Select(static (text, _) => ReadDocx(text))
            .WithTrackingName(Stages.Docs)
            .Collect()
            .Select(static (array, _) => new EquatableArray<DocxData>(array))
            .WithTrackingName(Stages.DocsCollected);

        var markdowns = context.AdditionalTextsProvider
            .Where(static _ => IsMarkdownPath(_.Path))
            .Select(static (text, cancel) => ReadMarkdown(text, cancel))
            .WithTrackingName(Stages.Markdowns)
            .Collect()
            .Select(static (array, _) => new EquatableArray<MarkdownData>(array))
            .WithTrackingName(Stages.MarkdownsCollected);

        var combined = targets
            .Combine(docs)
            .Combine(markdowns)
            .WithTrackingName(Stages.Combined);

        context.RegisterSourceOutput(
            combined,
            static (productionContext, tuple) =>
            {
                var targetInfos = tuple.Left.Left;
                var docData = tuple.Left.Right;
                var markdownData = tuple.Right;
                foreach (var target in targetInfos)
                {
                    Process(productionContext, target, docData, markdownData);
                }
            });
    }

    static bool IsMarkdownPath(string path) =>
        path.EndsWith(".md", StringComparison.OrdinalIgnoreCase);

    static TargetInfo? ExtractTarget(GeneratorAttributeSyntaxContext context, Cancel cancel)
    {
        if (context.TargetSymbol is not INamedTypeSymbol typeSymbol)
        {
            return null;
        }

        var attribute = context.Attributes.FirstOrDefault();
        if (attribute == null ||
            attribute.ConstructorArguments.Length < 1)
        {
            return null;
        }

        if (attribute.ConstructorArguments[0].Value is not string path)
        {
            return null;
        }

        var syntaxReference = attribute.ApplicationSyntaxReference;
        var rawLocation = syntaxReference == null
            ? Location.None
            : Location.Create(
                syntaxReference.SyntaxTree,
                syntaxReference.Span);

        var declaringNamespace = typeSymbol.ContainingNamespace.IsGlobalNamespace
            ? null
            : typeSymbol.ContainingNamespace.ToDisplayString();

        var enclosingResult = BuildEnclosingChain(typeSymbol);

        // The attribute target IS the model — there is no separate "marker / template" class.
        // ModelFullyQualifiedName / ModelSimpleName therefore describe the decorated class itself.
        var excelsiorTableType = context.SemanticModel.Compilation
            .GetTypeByMetadataName(ShapeBuilder.ExcelsiorTableAttributeFullName);
        var shape = ShapeBuilder.Build(typeSymbol, excelsiorTableType, cancel);

        return new(
            declaringNamespace,
            typeSymbol.Name,
            GetTypeKindKeyword(typeSymbol),
            new(enclosingResult.Chain.ToImmutableArray()),
            typeSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
            typeSymbol.Name,
            path,
            EquatableLocation.From(rawLocation),
            shape,
            enclosingResult.Error);
    }

    static (List<EnclosingType> Chain, string? Error) BuildEnclosingChain(INamedTypeSymbol typeSymbol)
    {
        var stack = new List<EnclosingType>();
        for (var current = typeSymbol.ContainingType; current != null; current = current.ContainingType)
        {
            // Every enclosing type must be `partial` — the SG emits the registration helper
            // wrapped in `partial {kind} {name} { ... }` declarations, and a non-partial
            // enclosing declaration would conflict with the user's existing one (CS0260).
            if (!IsPartial(current))
            {
                return (stack, current.Name);
            }

            stack.Add(new(current.Name, GetTypeKindKeyword(current)));
        }

        // ContainingType walks innermost → outermost; flip so emission can write outermost first.
        stack.Reverse();
        return (stack, null);
    }

    static bool IsPartial(INamedTypeSymbol typeSymbol)
    {
        foreach (var reference in typeSymbol.DeclaringSyntaxReferences)
        {
            if (reference.GetSyntax() is TypeDeclarationSyntax declaration &&
                declaration.Modifiers.Any(SyntaxKind.PartialKeyword))
            {
                return true;
            }
        }

        return false;
    }

    static string GetTypeKindKeyword(INamedTypeSymbol typeSymbol)
    {
        if (typeSymbol.IsRecord)
        {
            return typeSymbol.TypeKind == TypeKind.Struct ? "record struct" : "record";
        }

        return typeSymbol.TypeKind == TypeKind.Struct ? "struct" : "class";
    }

    static DocxData ReadDocx(AdditionalText text)
    {
        try
        {
            var paragraphs = DocxArchiveReader.ReadParagraphTexts(text.Path);
            return new(text.Path, new(paragraphs.ToImmutableArray()), null);
        }
        catch (Exception exception)
        {
            return new(text.Path, EquatableArray<string>.Empty, exception.Message);
        }
    }

    static MarkdownData ReadMarkdown(AdditionalText text, Cancel cancel)
    {
        try
        {
            // AdditionalText.GetText is the canonical Roslyn entry point — it handles encoding
            // detection and lets the SDK reuse any cached SourceText. Direct File.IO is banned
            // for analyzers (RS1035), so a null return here means the AdditionalText doesn't
            // back to a readable source; treat that as a read error.
            var sourceText = text.GetText(cancel);
            if (sourceText == null)
            {
                return new(text.Path, string.Empty, "AdditionalText returned no SourceText");
            }

            return new(text.Path, sourceText.ToString(), null);
        }
        catch (Exception exception)
        {
            return new(text.Path, string.Empty, exception.Message);
        }
    }

    static void Process(
        SourceProductionContext context,
        TargetInfo target,
        EquatableArray<DocxData> docs,
        EquatableArray<MarkdownData> markdowns)
    {
        var location = target.Location.ToLocation();

        if (target.ExtractError != null)
        {
            // PARCH011: an enclosing type isn't partial. Skip both validation and registration —
            // template tokens may still be valid, but emitting the registration helper into a
            // namespace-scope partial would land it in the wrong type.
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.EnclosingTypeNotPartial,
                    location,
                    target.DeclaringName,
                    target.ExtractError));
            return;
        }

        if (IsMarkdownPath(target.TemplatePath))
        {
            ProcessMarkdown(context, target, location, markdowns);
            return;
        }

        ProcessDocx(context, target, location, docs);
    }

    static void ProcessDocx(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        EquatableArray<DocxData> docs)
    {
        var normalized = target.TemplatePath.Replace('\\', '/');

        DocxData? matched = null;
        foreach (var doc in docs)
        {
            if (doc.Path.Replace('\\', '/')
                .EndsWith(normalized, StringComparison.OrdinalIgnoreCase))
            {
                matched = doc;
                break;
            }
        }

        if (matched == null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateFileMissing,
                    location,
                    target.TemplatePath));
            return;
        }

        if (matched.ReadError != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    target.TemplatePath,
                    matched.ReadError));
            return;
        }

        var tokens = TokenScanner.Scan(matched.Paragraphs);
        ValidateTokens(context, target, tokens, location);

        EmitRegistration(context, target, GenerateDocxRegistration(target));
    }

    static void ProcessMarkdown(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        EquatableArray<MarkdownData> markdowns)
    {
        var normalized = target.TemplatePath.Replace('\\', '/');

        MarkdownData? matched = null;
        foreach (var md in markdowns)
        {
            if (md.Path.Replace('\\', '/')
                .EndsWith(normalized, StringComparison.OrdinalIgnoreCase))
            {
                matched = md;
                break;
            }
        }

        if (matched == null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateFileMissing,
                    location,
                    target.TemplatePath));
            return;
        }

        if (matched.ReadError != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    target.TemplatePath,
                    matched.ReadError));
            return;
        }

        if (!markdownParser.TryParse(matched.Text, out var template, out var error))
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    target.TemplatePath,
                    $"Failed to parse markdown as a liquid template: {error}"));
            return;
        }

        MarkdownValidator.Validate(context, target, location, template);

        EmitRegistration(context, target, GenerateMarkdownRegistration(target));
    }

    static readonly FluidParser markdownParser = new();

    static void EmitRegistration(SourceProductionContext context, TargetInfo target, string source)
    {
        // The hint name must be unique across all targets in the compilation. Simple name alone
        // collides when two models share `DeclaringName` (e.g. `Outer1.Info` and `Outer2.Info`),
        // so include the namespace AND every enclosing-type name as part of the prefix.
        var builder = new StringBuilder();
        if (target.DeclaringNamespace != null)
        {
            builder.Append(target.DeclaringNamespace).Append('_');
        }

        foreach (var enclosing in target.EnclosingTypes)
        {
            builder.Append(enclosing.Name).Append('_');
        }

        builder.Append(target.DeclaringName);
        var hintPrefix = builder.ToString().Replace('.', '_');
        context.AddSource(
            $"{hintPrefix}_ParchmentModel.g.cs",
            SourceText.From(source, Encoding.UTF8));
    }

    static void ValidateTokens(
        SourceProductionContext context,
        TargetInfo target,
        IReadOnlyList<Token> tokens,
        Location location)
    {
        var scope = new Dictionary<string, string>(StringComparer.Ordinal);
        var loopStack = new Stack<string>();

        foreach (var token in tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.Substitution:
                    ValidateReferences(context, target, location, token.References, scope, token.Source);
                    ValidateExcelsiorToken(context, target, location, token, scope);
                    ValidateFormatToken(context, target, location, token, scope);
                    break;

                case TokenKind.ForOpen:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                target.TemplatePath,
                                token.Source));
                    }

                    if (token.LoopVariable == null ||
                        token.References.Count == 0)
                    {
                        break;
                    }

                    var sourceFqn = ShapeResolver.Resolve(target.Shape, token.References[0], scope);
                    if (sourceFqn == null)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MissingMember,
                                location,
                                target.TemplatePath,
                                token.Source,
                                string.Join('.', token.References[0]),
                                target.ModelSimpleName));
                        break;
                    }

                    var elementFqn = ShapeResolver.GetElementType(target.Shape, sourceFqn);
                    if (elementFqn == null)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.LoopSourceNotEnumerable,
                                location,
                                target.TemplatePath,
                                token.Source));
                        break;
                    }

                    scope[token.LoopVariable] = elementFqn;
                    loopStack.Push(token.LoopVariable);
                    break;

                case TokenKind.ForClose:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                target.TemplatePath,
                                token.Source));
                    }

                    if (loopStack.Count > 0)
                    {
                        scope.Remove(loopStack.Pop());
                    }

                    break;

                case TokenKind.IfOpen:
                case TokenKind.ElsIf:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                target.TemplatePath,
                                token.Source));
                    }

                    ValidateReferences(context, target, location, token.References, scope, token.Source);
                    break;

                case TokenKind.Else:
                case TokenKind.IfClose:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                target.TemplatePath,
                                token.Source));
                    }

                    break;

                case TokenKind.UnknownBlock:
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.UnsupportedBlockTag,
                            location,
                            target.TemplatePath,
                            token.Source));
                    break;
            }
        }
    }

    static void ValidateExcelsiorToken(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        Token token,
        Dictionary<string, string> scope)
    {
        // Only substitution tokens whose first identifier path resolves to an [ExcelsiorTable]
        // property need these extra checks. Everything else is handled by normal reference
        // validation (PARCH001/etc) or flows through the standard runtime substitution path.
        if (token.References.Count == 0)
        {
            return;
        }

        if (!ShapeResolver.IsExcelsiorTableMember(target.Shape, token.References[0], scope))
        {
            return;
        }

        if (token.HasOtherContent)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.ExcelsiorTokenNotAlone,
                    location,
                    target.TemplatePath,
                    token.Source));
        }

        if (!token.IsPlainIdentifier)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.ExcelsiorTokenNotPlainIdentifier,
                    location,
                    target.TemplatePath,
                    token.Source));
        }
    }

    static void ValidateFormatToken(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        Token token,
        Dictionary<string, string> scope)
    {
        if (token.References.Count == 0)
        {
            return;
        }

        var member = ShapeResolver.ResolveMember(target.Shape, token.References[0], scope);
        if (member is null or
            {
                IsHtml: false,
                IsMarkdown: false
            })
        {
            return;
        }

        if (token.IsPlainIdentifier)
        {
            return;
        }

        context.ReportDiagnostic(
            Diagnostic.Create(
                Diagnostics.FormatTokenNotPlainIdentifier,
                location,
                target.TemplatePath,
                token.Source));
    }

    static void ValidateReferences(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        List<List<string>> references,
        IReadOnlyDictionary<string, string> scope,
        string tokenSource)
    {
        foreach (var reference in references)
        {
            var resolved = ShapeResolver.Resolve(target.Shape, reference, scope);
            if (resolved != null)
            {
                continue;
            }

            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.MissingMember,
                    location,
                    target.TemplatePath,
                    tokenSource,
                    string.Join('.', reference),
                    target.ModelSimpleName));
        }
    }

    static string GenerateDocxRegistration(TargetInfo target)
    {
        var templatePath = target.TemplatePath.Replace("\\", @"\\");
        var body =
            $$"""
              public static string TemplatePath => "{{templatePath}}";
              public static string TemplateName => "{{target.DeclaringName}}";

              public static void RegisterWith(global::Parchment.TemplateStore store, string? basePath = null)
              {
                var path = basePath is null ? TemplatePath : global::System.IO.Path.Combine(basePath, TemplatePath);
                store.RegisterDocxTemplate<{{target.ModelFullyQualifiedName}}>(TemplateName, path);
              }
              """;

        return BuildPartialSource(target, body);
    }

    static string GenerateMarkdownRegistration(TargetInfo target)
    {
        var templatePath = target.TemplatePath.Replace("\\", @"\\");
        var body =
            $$"""
              public static string TemplatePath => "{{templatePath}}";
              public static string TemplateName => "{{target.DeclaringName}}";

              public static void RegisterWith(global::Parchment.TemplateStore store, string? basePath = null, global::System.IO.Stream? styleSource = null)
              {
                var path = basePath is null ? TemplatePath : global::System.IO.Path.Combine(basePath, TemplatePath);
                var markdown = global::System.IO.File.ReadAllText(path);
                store.RegisterMarkdownTemplate<{{target.ModelFullyQualifiedName}}>(TemplateName, markdown, styleSource);
              }
              """;

        return BuildPartialSource(target, body);
    }

    // Wraps `body` in `partial {kind} {name} { ... }` declarations: namespace (if any), then
    // each enclosing type outermost-first, then the target itself. Indentation isn't strictly
    // necessary for correctness but keeps the generated source readable in obj/.../generated.
    static string BuildPartialSource(TargetInfo target, string body)
    {
        var builder = new StringBuilder(
            """
            // <auto-generated />
            #nullable enable

            """);

        if (target.DeclaringNamespace != null)
        {
            builder.AppendLine($"namespace {target.DeclaringNamespace};");
        }

        var depth = 0;
        foreach (var enclosing in target.EnclosingTypes)
        {
            builder.Indent(depth).AppendLine($"partial {enclosing.Kind} {enclosing.Name}");
            builder.Indent(depth).AppendLine("{");
            depth++;
        }

        builder.Indent(depth).AppendLine($"partial {target.DeclaringKind} {target.DeclaringName}");
        builder.Indent(depth).AppendLine("{");
        foreach (var line in body.Split('\n'))
        {
            var trimmed = line.TrimEnd('\r');
            if (trimmed.Length == 0)
            {
                builder.AppendLine();
            }
            else
            {
                builder.Indent(depth + 1).AppendLine(trimmed);
            }
        }

        builder.Indent(depth).AppendLine("}");

        for (var i = depth - 1; i >= 0; i--)
        {
            builder.Indent(i).AppendLine("}");
        }

        return builder.ToString();
    }

    public static class Stages
    {
        public const string Targets = "Parchment_Targets";
        public const string TargetsCollected = "Parchment_TargetsCollected";
        public const string Docs = "Parchment_Docs";
        public const string DocsCollected = "Parchment_DocsCollected";
        public const string Markdowns = "Parchment_Markdowns";
        public const string MarkdownsCollected = "Parchment_MarkdownsCollected";
        public const string Combined = "Parchment_Combined";
    }
}
