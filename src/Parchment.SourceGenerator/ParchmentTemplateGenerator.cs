[Generator]
public sealed class ParchmentTemplateGenerator :
    IIncrementalGenerator
{
    const string attributeFullName = "Parchment.ParchmentModelAttribute";
    const string bindableAttributeFullName = "Parchment.ParchmentBindableAttribute";

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

        // Filtered to the marked files before anything reads them. An unmarked .md AdditionalFile
        // is some other tool's business, and parsing it as liquid to discover that would be work
        // repeated every build to reach the same answer.
        var docs = context.AdditionalTextsProvider
            .Where(static _ => TemplateConvention.IsDocx(_.Path))
            .Select(static (text, _) => ReadDocx(text))
            .WithTrackingName(Stages.Docs)
            .Collect()
            .Select(static (array, _) => new EquatableArray<DocxData>(array))
            .WithTrackingName(Stages.DocsCollected);

        var markdowns = context.AdditionalTextsProvider
            .Where(static _ => TemplateConvention.IsMarkdown(_.Path))
            .Select(static (text, cancel) => ReadMarkdown(text, cancel))
            .WithTrackingName(Stages.Markdowns)
            .Collect()
            .Select(static (array, _) => new EquatableArray<MarkdownData>(array))
            .WithTrackingName(Stages.MarkdownsCollected);

        var dotxes = context.AdditionalTextsProvider
            .Where(static _ => TemplateConvention.IsStyleDoc(_.Path))
            .Select(static (text, _) => ReadDotx(text))
            .WithTrackingName(Stages.Dotxes)
            .Collect()
            .Select(static (array, _) => new EquatableArray<DotxData>(array))
            .WithTrackingName(Stages.DotxesCollected);

        // The project directory shortens a template's absolute path to a project-relative one for
        // diagnostic messages.
        var projectDirectory = context.AnalyzerConfigOptionsProvider
            .Select(static (options, _) =>
                options.GlobalOptions.TryGetValue("build_property.ProjectDir", out var directory)
                    ? directory
                    : null);

        var combined = targets
            .Combine(docs)
            .Combine(markdowns)
            .Combine(dotxes)
            .Combine(projectDirectory)
            .WithTrackingName(Stages.Combined);

        context.RegisterSourceOutput(
            combined,
            static (productionContext, tuple) =>
            {
                var targetInfos = tuple.Left.Left.Left.Left;
                var docData = tuple.Left.Left.Left.Right;
                var markdownData = tuple.Left.Left.Right;
                var dotxData = tuple.Left.Right;
                var projectDir = tuple.Right;
                foreach (var target in targetInfos)
                {
                    Process(productionContext, target, docData, markdownData, dotxData, projectDir);
                }
            });

        // [ParchmentBindable]: accessors-only emission for models registered by hand against
        // templates the generator cannot see. No template inputs to combine, so each target is
        // its own cacheable unit.
        var bindables = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                bindableAttributeFullName,
                static (node, _) => node is ClassDeclarationSyntax or RecordDeclarationSyntax,
                ExtractBindableTarget)
            .Where(static _ => _ != null)
            .Select(static (target, _) => target!)
            .WithTrackingName(Stages.Bindables);

        context.RegisterSourceOutput(
            bindables,
            static (productionContext, target) => ProcessBindable(productionContext, target));
    }

    static TargetInfo? ExtractBindableTarget(GeneratorAttributeSyntaxContext context, Cancel cancel)
    {
        // [ParchmentModel] already emits the accessors (together with the embedded template); a
        // second module initializer here would collide with it, so the template attribute wins.
        if (context.TargetSymbol is INamedTypeSymbol symbol &&
            symbol.GetAttributes().Any(_ => _.AttributeClass?.ToDisplayString() == "Parchment.ParchmentModelAttribute"))
        {
            return null;
        }

        return ExtractTarget(context, cancel);
    }

    static void ProcessBindable(SourceProductionContext context, TargetInfo target)
    {
        var location = target.Location.ToLocation();

        if (target.ExtractError != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.EnclosingTypeNotPartial,
                    location,
                    target.DeclaringName,
                    target.ExtractError));
            return;
        }

        // The template kind is unknown at compile time, so the shape rules that a docx
        // registration would need are enforced here — an invalid editable member is an authoring
        // mistake regardless of which template it later meets.
        ValidateEditableShape(context, target, location);

        EmitRegistration(context, target, GenerateBindableRegistration(target), "ParchmentBindable");
    }

    static string GenerateBindableRegistration(TargetInfo target)
    {
        var (fieldsBlock, registrationsBlock) = PrepareAccessors(target);
        // The type name is folded into the method name so a base and a derived model can both be
        // [ParchmentBindable] without the derived initializer hiding the base one (CS0108).
        var body =
            $$"""
              {{fieldsBlock}}[global::System.Runtime.CompilerServices.ModuleInitializer]
              internal static void InitializeParchmentBindable{{target.DeclaringName}}()
              {
              {{registrationsBlock}}}
              """;

        return BuildPartialSource(target, body);
    }

    static TargetInfo? ExtractTarget(GeneratorAttributeSyntaxContext context, Cancel cancel)
    {
        if (context.TargetSymbol is not INamedTypeSymbol typeSymbol)
        {
            return null;
        }

        var attribute = context.Attributes.FirstOrDefault();
        if (attribute == null)
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
        // ModelFullyQualifiedName / ModelDisplayName therefore describe the decorated class itself.
        // DisplayName joins enclosing types dotted-style so PARCH001 messages for nested models
        // (e.g. `XxxGenerator.Info` patterns) disambiguate from sibling `Info` types.
        var displayName = enclosingResult.Chain.Count == 0
            ? typeSymbol.Name
            : string.Join('.', enclosingResult.Chain.Select(_ => _.Name)) + '.' + typeSymbol.Name;

        var excelsiorTableType = context.SemanticModel.Compilation
            .GetTypeByMetadataName(ShapeBuilder.ExcelsiorTableAttributeFullName);
        var editableFieldType = context.SemanticModel.Compilation
            .GetTypeByMetadataName(ShapeBuilder.EditableFieldAttributeFullName);
        var shape = ShapeBuilder.Build(typeSymbol, excelsiorTableType, editableFieldType, cancel);

        var protection = ProtectionMode.WhenEditable;
        foreach (var named in attribute.NamedArguments)
        {
            // The named argument arrives as the enum's underlying int — the SG cannot reference
            // Parchment.dll's ProtectionMode type.
            if (named is { Key: "Protection", Value.Value: int protectionValue })
            {
                protection = (ProtectionMode)protectionValue;
            }
        }

        return new(
            declaringNamespace,
            typeSymbol.Name,
            GetTypeKindKeyword(typeSymbol),
            new(enclosingResult.Chain.ToImmutableArray()),
            typeSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
            displayName,
            protection,
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

    static string GetTypeKindKeyword(INamedTypeSymbol type)
    {
        if (type.IsRecord)
        {
            return type.TypeKind == TypeKind.Struct ? "record struct" : "record";
        }

        return type.TypeKind == TypeKind.Struct ? "struct" : "class";
    }

    // Where the file sits relative to the project — the shortest way to name it in a diagnostic.
    // Falls back to the file name when the project directory is unknown or the file lives outside
    // it.
    /// <summary>
    /// The templates the generator was handed, for a model that matched none of them.
    /// </summary>
    /// <remarks>
    /// Which of the two failures this is turns entirely on that list. Names in it mean the files
    /// arrived and the model is looking for one nothing is called — usually a near-miss sitting
    /// right there in the message. An empty list means no template reached the generator at all,
    /// which is a build or IDE problem rather than a naming one, and no amount of renaming fixes
    /// it. The message cannot tell them apart; the list can.
    /// </remarks>
    static string DescribeTemplatesSeen(
        EquatableArray<DocxData> docs,
        EquatableArray<MarkdownData> markdowns,
        string? projectDirectory)
    {
        var paths = docs.Select(static _ => _.Path)
            .Concat(markdowns.Select(static _ => _.Path))
            .Select(_ => DisplayPath(_, projectDirectory))
            // Ordered rather than left in AdditionalFiles order: the message should read the same
            // way whatever order the item groups happened to be evaluated in.
            .OrderBy(static _ => _, StringComparer.Ordinal)
            .ToList();

        if (paths.Count == 0)
        {
            return "No templates reached the generator at all.";
        }

        if (paths.Count > maxTemplatesListed)
        {
            var listed = string.Join(", ", paths.Take(maxTemplatesListed));
            return $"Templates seen: {listed} (+{paths.Count - maxTemplatesListed} more).";
        }

        return $"Templates seen: {string.Join(", ", paths)}.";
    }

    // A project with hundreds of templates would otherwise turn one missing name into a diagnostic
    // nobody reads to the end of.
    const int maxTemplatesListed = 10;

    static string DisplayPath(string fullPath, string? projectDirectory)
    {
        if (projectDirectory != null)
        {
            var project = TemplateConvention.Normalize(projectDirectory);
            var file = TemplateConvention.Normalize(fullPath);
            if (file.StartsWith($"{project}/", StringComparison.OrdinalIgnoreCase))
            {
                return file.Substring(project.Length + 1);
            }
        }

        return Path.GetFileName(fullPath);
    }

    static DocxData ReadDocx(AdditionalText text)
    {
        try
        {
            var result = DocxArchiveReader.Read(text.Path);
            return new(
                text.Path,
                new(result.Paragraphs.ToImmutableArray()),
                new(result.BodyParagraphs.ToImmutableArray()),
                result.HasRemovePersonalInformation,
                Convert.ToBase64String(ReadBytes(text.Path)),
                null);
        }
        catch (Exception exception)
        {
            return new(text.Path, EquatableArray<string>.Empty, EquatableArray<string>.Empty, false, string.Empty, exception.Message);
        }
    }

    static DotxData ReadDotx(AdditionalText text)
    {
        try
        {
            return new(text.Path, Convert.ToBase64String(ReadBytes(text.Path)), null);
        }
        catch (Exception exception)
        {
            return new(text.Path, string.Empty, exception.Message);
        }
    }

    // RS1035 steers analyzers to AdditionalText, but an AdditionalText has no binary read surface
    // — GetText() returns null for binary content — so raw bytes come off the stream directly,
    // the same loophole DocxArchiveReader relies on for ZipFile.
#pragma warning disable RS1035
    static byte[] ReadBytes(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
#pragma warning restore RS1035

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
        EquatableArray<MarkdownData> markdowns,
        EquatableArray<DotxData> dotxes,
        string? projectDirectory)
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

        // Convention: the template is the AdditionalFile named after the type. Nothing matched is
        // PARCH004; more than one — a docx and an md, or namesakes in different folders — is
        // PARCH020, reported rather than resolved by preference, since silently picking one would
        // bind the model to a template its author was not looking at.
        var docxCandidates = TemplateConvention.MatchByTypeName(docs, static _ => _.Path, target.DeclaringName);
        var markdownCandidates = TemplateConvention.MatchByTypeName(markdowns, static _ => _.Path, target.DeclaringName);

        var total = docxCandidates.Count + markdownCandidates.Count;
        if (total == 0)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateFileMissing,
                    location,
                    target.DeclaringName,
                    DescribeTemplatesSeen(docs, markdowns, projectDirectory)));
            return;
        }

        if (total > 1)
        {
            var paths = docxCandidates.Select(static _ => _.Path)
                .Concat(markdownCandidates.Select(static _ => _.Path))
                .Select(_ => DisplayPath(_, projectDirectory));
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.AmbiguousTemplate,
                    location,
                    target.DeclaringName,
                    string.Join(", ", paths)));
            return;
        }

        if (markdownCandidates.Count == 1)
        {
            ProcessMarkdown(context, target, location, markdownCandidates[0], dotxes, projectDirectory);
            return;
        }

        ProcessDocx(context, target, location, docxCandidates[0], projectDirectory);
    }

    static void ProcessDocx(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        DocxData matched,
        string? projectDirectory)
    {
        var templatePath = DisplayPath(matched.Path, projectDirectory);
        if (matched.ReadError != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    templatePath,
                    matched.ReadError));
            return;
        }

        if (!matched.HasRemovePersonalInformation)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.MissingRemovePersonalInformation,
                    location,
                    templatePath));
        }

        var tokens = TokenScanner.Scan(matched.Paragraphs);
        ValidateTokens(context, target, templatePath, tokens, location);

        // Editable fields are docx-only (like Excelsior / Format dispatch), so the shape rules
        // fire here rather than in Process — the runtime lockstep is EditableMap.Build throwing
        // at RegisterDocxTemplate regardless of whether a token references the member.
        ValidateEditableShape(context, target, location);
        if (HasEditableMembers(target.Shape))
        {
            // Token rules are body-scoped: the runtime dispatches editable fields only in the
            // document body, so a header/footer occurrence of the same member is a deliberate
            // read-only mirror, not a duplicate.
            var bodyTokens = TokenScanner.Scan(matched.BodyParagraphs);
            ValidateEditableTokens(context, target, templatePath, bodyTokens, location);
        }

        EmitRegistration(context, target, GenerateDocxRegistration(target, matched));
    }

    static bool HasEditableMembers(ModelShape shape)
    {
        foreach (var type in shape.Types)
        {
            foreach (var member in type.Members)
            {
                if (member is { IsEditable: true, IsStatic: false })
                {
                    return true;
                }
            }
        }

        return false;
    }

    // IsStringList is inferred from the member's type rather than written by hand, so it is not
    // reported — a static IEnumerable<string> is not a mistake the way a hand-written attribute is.
    static string? IgnoredStaticAttribute(MemberEntry member)
    {
        if (member.IsExcelsiorTable)
        {
            return "ExcelsiorTable";
        }

        if (member.IsHtml)
        {
            return "Html";
        }

        if (member.IsMarkdown)
        {
            return "Markdown";
        }

        if (member.IsEditable)
        {
            return "EditableField";
        }

        return null;
    }

    static void ValidateEditableShape(
        SourceProductionContext context,
        TargetInfo target,
        Location location)
    {
        foreach (var type in target.Shape.Types)
        {
            foreach (var member in type.Members)
            {
                var memberDisplay = $"{Display(type.TypeFullyQualifiedName)}.{member.Name}";

                // A render attribute on a static member is a no-op: the per-template maps that
                // dispatch it walk instance members only, so the value renders as plain text and
                // the attribute is dropped. That is silent at runtime, hence the warning.
                if (member.IsStatic)
                {
                    var ignored = IgnoredStaticAttribute(member);
                    if (ignored != null)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.RenderAttributeOnStaticMember,
                                location,
                                memberDisplay,
                                ignored));
                    }

                    continue;
                }

                if (member.HasFormatConflict)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.ConflictingFormatMarkers,
                            location,
                            target.ModelDisplayName,
                            memberDisplay));
                }

                if (member.FormatMarkerOnTokenValue)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.FormatMarkerOnTokenValue,
                            location,
                            target.ModelDisplayName,
                            memberDisplay));
                }

                if (!member.IsEditable)
                {
                    continue;
                }

                // [Html] + [EditableField] is supported (renders an editable rich-content block that
                // extracts back to HTML). [ExcelsiorTable] and [Markdown] remain conflicts.
                if (member.IsExcelsiorTable ||
                    member.IsMarkdown)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.EditableConflictingAttribute,
                            location,
                            target.ModelDisplayName,
                            memberDisplay));
                    continue;
                }

                // An [EditableField] collection of a POCO element type renders as a repeating section
                // (EditableKind stays null). Extraction rebuilds the list from the repeated items, so
                // the collection needs a setter and its element type must be constructable, carry at
                // least one editable member, and not nest a further editable collection.
                if (member.EditableCollectionElementFqn != null)
                {
                    if (!member.HasUsableSetter)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.EditableNoSetter,
                                location,
                                target.ModelDisplayName,
                                memberDisplay));
                    }

                    ValidateEditableCollection(context, target, location, memberDisplay, member);
                    continue;
                }

                if (member.EditableKind == null)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.EditableUnsupportedType,
                            location,
                            target.ModelDisplayName,
                            memberDisplay,
                            Display(member.TypeFullyQualifiedName)));
                    continue;
                }

                if (!member.HasUsableSetter)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.EditableNoSetter,
                            location,
                            target.ModelDisplayName,
                            memberDisplay));
                }
            }
        }
    }

    /// <summary>
    /// The collection-shape half of the editable rules. Lockstep with the emission-side skip in
    /// <c>AccessorEmission.WalkForMaps</c>: a shape reported here is not emitted, so the generated
    /// source stays compilable while the diagnostic fails the build.
    /// </summary>
    static void ValidateEditableCollection(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        string memberDisplay,
        MemberEntry member)
    {
        TypeEntry? element = null;
        foreach (var type in target.Shape.Types)
        {
            if (type.TypeFullyQualifiedName == member.EditableCollectionElementFqn)
            {
                element = type;
                break;
            }
        }

        if (element == null)
        {
            return;
        }

        var elementDisplay = Display(element.TypeFullyQualifiedName);
        if (!element.HasParameterlessCtor)
        {
            ReportCollectionInvalid(
                context,
                target,
                location,
                memberDisplay,
                $"its element type '{elementDisplay}' has no public parameterless constructor — extraction rebuilds the list, so elements must be constructable");
        }

        var hasEditable = false;
        foreach (var elementMember in element.Members)
        {
            if (elementMember.EditableCollectionElementFqn != null)
            {
                ReportCollectionInvalid(
                    context,
                    target,
                    location,
                    memberDisplay,
                    $"its element type '{elementDisplay}' itself contains an editable collection — nested editable collections are not supported");
            }
            else if (elementMember is { IsEditable: true, IsStatic: false })
            {
                hasEditable = true;
            }
        }

        if (!hasEditable)
        {
            ReportCollectionInvalid(
                context,
                target,
                location,
                memberDisplay,
                $"its element type '{elementDisplay}' has no [EditableField] members — nothing to round-trip");
        }
    }

    static void ReportCollectionInvalid(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        string memberDisplay,
        string reason) =>
        context.ReportDiagnostic(
            Diagnostic.Create(
                Diagnostics.EditableCollectionInvalid,
                location,
                target.ModelDisplayName,
                memberDisplay,
                reason));

    static string Display(string fqn)
    {
        if (fqn.StartsWith("global::", StringComparison.Ordinal))
        {
            return fqn["global::".Length..];
        }

        return fqn;
    }

    /// <summary>
    /// Body-scoped editable-token rules. This pass tracks loop scope silently (no diagnostics —
    /// <see cref="ValidateTokens"/> already reported PARCH001/002/005 over all parts) so PARCH016
    /// / PARCH017 / PARCH018 don't double-report.
    /// </summary>
    static void ValidateEditableTokens(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        IReadOnlyList<Token> tokens,
        Location location)
    {
        var scope = new Dictionary<string, string>(StringComparer.Ordinal);
        var loopStack = new Stack<(string? Name, string? PriorBinding, bool IsEditableCollection)>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var token in tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.ForOpen:
                    string? bound = null;
                    string? prior = null;
                    var isEditableCollectionLoop = false;
                    if (token is { LoopVariable: not null, References.Count: > 0 })
                    {
                        isEditableCollectionLoop =
                            ShapeResolver.TryResolveMember(target.Shape, token.References[0], scope, out var sourceMember) &&
                            sourceMember.EditableCollectionElementFqn != null;

                        if (ShapeResolver.TryResolve(target.Shape, token.References[0], scope, out var sourceFqn) &&
                            ShapeResolver.TryGetElementType(target.Shape, sourceFqn, out var elementFqn))
                        {
                            bound = token.LoopVariable;
                            prior = scope.GetValueOrDefault(bound);
                            scope[bound] = elementFqn;
                        }
                    }

                    loopStack.Push((bound, prior, isEditableCollectionLoop));
                    break;

                case TokenKind.ForClose:
                    if (loopStack.Count > 0)
                    {
                        var (name, priorBinding, _) = loopStack.Pop();
                        if (name != null)
                        {
                            if (priorBinding == null)
                            {
                                scope.Remove(name);
                            }
                            else
                            {
                                scope[name] = priorBinding;
                            }
                        }
                    }

                    break;

                case TokenKind.Substitution:
                    if (token.References.Count == 0)
                    {
                        break;
                    }

                    if (!ShapeResolver.TryResolveMember(target.Shape, token.References[0], scope, out var member) ||
                        member is { IsEditable: false } or { IsStatic: true })
                    {
                        break;
                    }

                    // A loop over an editable collection turns its loop-variable tokens into
                    // repeating-section controls, so PARCH018 (editable token in a loop) does not apply.
                    var insideEditableCollection = false;
                    foreach (var frame in loopStack)
                    {
                        if (frame.IsEditableCollection)
                        {
                            insideEditableCollection = true;
                            break;
                        }
                    }

                    if (loopStack.Count > 0 &&
                        !insideEditableCollection)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.EditableTokenInLoop,
                                location,
                                templatePath,
                                token.Source));
                        break;
                    }

                    // Editable-collection loop-variable tokens are handled by the repeating-section
                    // render — they don't participate in the flat body's plain-identifier / duplicate
                    // checks (their tags are item-relative), so stop here.
                    if (insideEditableCollection)
                    {
                        break;
                    }

                    if (!token.IsPlainIdentifier)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.EditableTokenNotPlainIdentifier,
                                location,
                                templatePath,
                                token.Source));
                    }

                    if (!seen.Add(string.Join('.', token.References[0])))
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.EditableTokenDuplicated,
                                location,
                                templatePath,
                                token.Source));
                    }

                    break;
            }
        }
    }

    /// <summary>
    /// The markdown flow's PARCH007 / PARCH008, which the AST walk cannot reach.
    /// </summary>
    /// <remarks>
    /// An Excelsior token is rewritten to a marker before the source is parsed, so what disqualifies
    /// one is a property of the line it sits on rather than of the Fluid AST — whether the token is
    /// the whole line, and whether the expression is a bare path. Both are read off the source here,
    /// the same way <c>MarkdownExcelsiorTables</c> reads them at registration.
    ///
    /// Without this the mistake still gets caught, but not until the store materializes the template
    /// on the model's first render — a runtime failure for something the build can see.
    /// </remarks>
    static void ValidateMarkdownExcelsiorTokens(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        Location location,
        string markdown)
    {
        foreach (var rawLine in markdown.Split('\n'))
        {
            var line = rawLine.Trim();
            foreach (Match match in markdownToken.Matches(line))
            {
                var body = match.Groups["body"].Value.Trim();
                var end = body.IndexOfAny([' ', '|', '=', '<', '>', '!', '+', '-', '*', '/', '[', '(']);
                var path = (end < 0 ? body : body[..end]).Replace(" ", "").Replace("\t", "");
                if (path.Length == 0)
                {
                    continue;
                }

                // Root paths only, matching the map the render looks the table up in: a loop
                // variable's members are not reachable from the root and fall through to Fluid.
                if (!ShapeResolver.IsExcelsiorTableMember(target.Shape, path.Split('.'), EmptyScope))
                {
                    continue;
                }

                if (match.Value != line)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.ExcelsiorTokenNotAlone,
                            location,
                            templatePath,
                            match.Value));
                    continue;
                }

                if (body != path)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.ExcelsiorTokenNotPlainIdentifier,
                            location,
                            templatePath,
                            match.Value));
                }
            }
        }
    }

    static readonly Dictionary<string, string> EmptyScope = new();

    static readonly Regex markdownToken = new(@"\{\{(?<body>[^{}]*)\}\}", RegexOptions.Compiled);

    static void ProcessMarkdown(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        MarkdownData matched,
        EquatableArray<DotxData> dotxes,
        string? projectDirectory)
    {
        var templatePath = DisplayPath(matched.Path, projectDirectory);
        if (matched.ReadError != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    templatePath,
                    matched.ReadError));
            return;
        }

        if (!markdownParser.TryParse(matched.Text, out var template, out var error))
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    templatePath,
                    $"Failed to parse markdown as a liquid template: {error}"));
            return;
        }

        MarkdownValidator.Validate(context, target, templatePath, location, template);
        ValidateMarkdownExcelsiorTokens(context, target, templatePath, location, matched.Text);

        // The style source: the dotx named after the type when the project carries one, otherwise
        // the nearest parchment.dotx up the directory tree from the template. Neither is required
        // — without one the markdown renders against the built-in blank document.
        var styleCandidates = TemplateConvention.MatchByTypeName(dotxes, static _ => _.Path, target.DeclaringName);
        if (styleCandidates.Count > 1)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.AmbiguousStyleDoc,
                    location,
                    target.DeclaringName,
                    string.Join(", ", styleCandidates.Select(_ => DisplayPath(_.Path, projectDirectory)))));
            return;
        }

        var style = styleCandidates.Count == 1
            ? styleCandidates[0]
            : TemplateConvention.FindSharedStyleDoc(dotxes, matched.Path);

        if (style?.ReadError != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.TemplateReadError,
                    location,
                    DisplayPath(style.Path, projectDirectory),
                    style.ReadError));
            return;
        }

        EmitRegistration(context, target, GenerateMarkdownRegistration(target, matched, style));
    }

    static readonly FluidParser markdownParser = new();

    static void EmitRegistration(SourceProductionContext context, TargetInfo target, string source, string suffix = "ParchmentModel")
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
            $"{hintPrefix}_{suffix}.g.cs",
            SourceText.From(source, Encoding.UTF8));
    }

    static void ValidateTokens(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        IReadOnlyList<Token> tokens,
        Location location)
    {
        var scope = new Dictionary<string, string>(StringComparer.Ordinal);
        // Stack entries carry the prior binding (if any) so nested loops with the same variable
        // name can restore the outer binding on ForClose instead of removing the entry outright.
        // Mirrors the save/restore pattern in MarkdownValidator.WalkFor.
        var loopStack = new Stack<(string Name, string? PriorBinding)>();

        foreach (var token in tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.Substitution:
                    ValidateReferences(context, target, templatePath, location, token.References, scope, token.Source);
                    ValidateExcelsiorToken(context, target, templatePath, location, token, scope);
                    ValidateFormatToken(context, target, templatePath, location, token, scope);
                    break;

                case TokenKind.ForOpen:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                templatePath,
                                token.Source));
                    }

                    if (token.LoopVariable == null ||
                        token.References.Count == 0)
                    {
                        break;
                    }

                    if (!ShapeResolver.TryResolve(target.Shape, token.References[0], scope, out var sourceFqn))
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MissingMember,
                                location,
                                templatePath,
                                token.Source,
                                string.Join('.', token.References[0]),
                                target.ModelDisplayName));
                        break;
                    }

                    if (!ShapeResolver.TryGetElementType(target.Shape, sourceFqn, out var elementFqn))
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.LoopSourceNotEnumerable,
                                location,
                                templatePath,
                                token.Source));
                        break;
                    }

                    var hadPriorBinding = scope.TryGetValue(token.LoopVariable, out var priorBinding);
                    scope[token.LoopVariable] = elementFqn;
                    loopStack.Push((token.LoopVariable, hadPriorBinding ? priorBinding : null));
                    break;

                case TokenKind.ForClose:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                templatePath,
                                token.Source));
                    }

                    if (loopStack.Count > 0)
                    {
                        var (name, prior) = loopStack.Pop();
                        if (prior == null)
                        {
                            scope.Remove(name);
                        }
                        else
                        {
                            scope[name] = prior;
                        }
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
                                templatePath,
                                token.Source));
                    }

                    ValidateReferences(context, target, templatePath, location, token.References, scope, token.Source);
                    break;

                case TokenKind.Else:
                case TokenKind.IfClose:
                    if (token.HasOtherContent)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.MixedInlineBlockTag,
                                location,
                                templatePath,
                                token.Source));
                    }

                    break;

                case TokenKind.UnknownBlock:
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.UnsupportedBlockTag,
                            location,
                            templatePath,
                            token.Source));
                    break;
            }
        }
    }

    static void ValidateExcelsiorToken(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
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
                    templatePath,
                    token.Source));
        }

        if (!token.IsPlainIdentifier)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.ExcelsiorTokenNotPlainIdentifier,
                    location,
                    templatePath,
                    token.Source));
        }
    }

    static void ValidateFormatToken(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        Location location,
        Token token,
        Dictionary<string, string> scope)
    {
        if (token.References.Count == 0)
        {
            return;
        }

        if (!ShapeResolver.TryResolveMember(target.Shape, token.References[0], scope, out var member) ||
            member is
            {
                IsHtml: false,
                IsMarkdown: false
            })
        {
            return;
        }

        // Rooted at a loop variable, so the format marker cannot be honoured: the runtime resolves
        // one by walking the model from its root, which has no way to address the item the current
        // iteration is on. The value renders as text with its markup showing, and nothing else says
        // so — see MarkdownFormats and ScopeTreeRunner.TryResolveFormatted.
        if (scope.ContainsKey(token.References[0][0]))
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.FormatMarkerThroughLoopVariable,
                    location,
                    templatePath,
                    token.Source,
                    member.Name));
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
                templatePath,
                token.Source));
    }

    static void ValidateReferences(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        Location location,
        IReadOnlyList<IReadOnlyList<string>> references,
        IReadOnlyDictionary<string, string> scope,
        string tokenSource)
    {
        foreach (var reference in references)
        {
            if (ShapeResolver.TryResolve(target.Shape, reference, scope, out _))
            {
                continue;
            }

            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.MissingMember,
                    location,
                    templatePath,
                    tokenSource,
                    string.Join('.', reference),
                    target.ModelDisplayName));
        }
    }

    // Each accessor section is emitted on its own physical lines so BuildPartialSource's
    // line-by-line outer indent pass adds the right depth prefix to every line. The blocks
    // already carry inner indentation; outer pass tops them up by depth+1.
    static (string FieldsBlock, string RegistrationsBlock) PrepareAccessors(TargetInfo target)
    {
        var accessors = AccessorEmission.Emit(target.Shape, target.ModelFullyQualifiedName);
        var fieldsBlock = accessors == null ? "" : accessors.FieldsBlock + "\n\n";
        var registrationsBlock = accessors == null ? "" : accessors.RegistrationsBlock + "\n";
        return (fieldsBlock, registrationsBlock);
    }

    // The template travels inside the generated source — bytes as base64, markdown as a string
    // literal — and a module initializer hands it to GeneratedRegistration when the model's
    // assembly loads. There is nothing to deploy beside the assembly and nothing to read at
    // runtime; a TemplateStore materializes the definition on the model's first render.
    static string GenerateDocxRegistration(TargetInfo target, DocxData matched)
    {
        var (fieldsBlock, registrationsBlock) = PrepareAccessors(target);
        // Emit the protection argument only when it deviates from the default, so registrations
        // for the common case stay minimal.
        var protection = target.Protection == ProtectionMode.WhenEditable
            ? ""
            : $", global::Parchment.ProtectionMode.{target.Protection}";
        var body =
            $$"""
              {{fieldsBlock}}[global::System.Runtime.CompilerServices.ModuleInitializer]
              internal static void InitializeParchmentTemplate()
              {
              {{registrationsBlock}}  global::Parchment.Generated.GeneratedRegistration.RegisterDocxTemplate(
                  typeof({{target.ModelFullyQualifiedName}}),
                  global::System.Convert.FromBase64String("{{matched.ContentBase64}}"){{protection}});
              }
              """;

        return BuildPartialSource(target, body);
    }

    static string GenerateMarkdownRegistration(TargetInfo target, MarkdownData matched, DotxData? style)
    {
        var (fieldsBlock, registrationsBlock) = PrepareAccessors(target);
        var markdown = SymbolDisplay.FormatLiteral(matched.Text, quote: true);
        var styleArgument = style == null
            ? ""
            : $",\n    global::System.Convert.FromBase64String(\"{style.ContentBase64}\")";
        var body =
            $$"""
              {{fieldsBlock}}[global::System.Runtime.CompilerServices.ModuleInitializer]
              internal static void InitializeParchmentTemplate()
              {
              {{registrationsBlock}}  global::Parchment.Generated.GeneratedRegistration.RegisterMarkdownTemplate(
                  typeof({{target.ModelFullyQualifiedName}}),
                  {{markdown}}{{styleArgument}});
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

            using System.Collections.Generic;

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

        // Everything emitted here lives in a nested type of its own rather than directly in the
        // model, so [ExcludeFromCodeCoverage] can sit on it. The accessor tables are one
        // DelegateAccessor lambda per member of every type reachable from the model, and only the
        // members a template actually binds are ever invoked — counted as covered code they measure
        // the shape of the model graph rather than how well the consumer is tested, and drag the
        // total down accordingly. The attribute cannot go on the model's own partial: that would
        // exclude the consumer's own members too.
        //
        // The type name is folded into the nested type's name for the reason the bindable
        // initializer folds it into the method's: a derived model would otherwise hide the base's
        // nested type (CS0108). Internal rather than private because a module initializer has to be
        // reachable from the module, which a private nested type is not.
        builder.Indent(depth + 1).AppendLine("[global::System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]");
        builder.Indent(depth + 1).AppendLine($"internal static class ParchmentGenerated{target.DeclaringName}");
        builder.Indent(depth + 1).AppendLine("{");
        foreach (var line in body.Split('\n'))
        {
            var trimmed = line.TrimEnd('\r');
            if (trimmed.Length == 0)
            {
                builder.AppendLine();
            }
            else
            {
                builder.Indent(depth + 2).AppendLine(trimmed);
            }
        }

        builder.Indent(depth + 1).AppendLine("}");
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
        public const string Dotxes = "Parchment_Dotxes";
        public const string DotxesCollected = "Parchment_DotxesCollected";
        public const string Combined = "Parchment_Combined";
        public const string Bindables = "Parchment_Bindables";
    }
}
