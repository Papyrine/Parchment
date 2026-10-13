[Generator]
public sealed class ParchmentTemplateGenerator :
    IIncrementalGenerator
{
    const string attributeFullName = "Parchment.ParchmentModelAttribute";
    const string bindableAttributeFullName = "Parchment.ParchmentBindableAttribute";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // ForAttributeWithMetadataName re-runs ExtractTarget for every attributed class on each
        // new compilation, so the shape always reflects the current symbols: editing a model
        // class in a separate file (e.g. adding/removing a property on Customer) changes the
        // extracted TargetInfo and re-validates the templates that reference it. What keeps the
        // pipeline cheap is that TargetInfo is primitive-only and compared by value — an edit
        // that leaves the shape alone yields an equal TargetInfo and everything downstream
        // stays cached.
        var targets = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                attributeFullName,
                static (node, _) => node is ClassDeclarationSyntax or RecordDeclarationSyntax,
                static (context, cancel) => Locate(context, ExtractTarget(context, cancel)))
            .Where(static _ => _ != null)
            .Select(static (target, _) => target!)
            .WithTrackingName(Stages.Targets)
            .Collect()
            .Select(static (array, _) => new EquatableArray<LocatedTarget>(array))
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
            {
                if (options.GlobalOptions.TryGetValue("build_property.ProjectDir", out var directory))
                {
                    return directory;
                }

                return null;
            });

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
                // A shared parchment.dotx is every model's style source in its folder tree, so an
                // unreadable one is reported against the file once, not once per model.
                var reportedStyleDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var target in targetInfos)
                {
                    Process(productionContext, target, docData, markdownData, dotxData, projectDir, reportedStyleDocs);
                }
            });

        // [ParchmentBindable]: accessors-only emission for models registered by hand against
        // templates the generator cannot see. No template inputs to combine, so each target is
        // its own cacheable unit.
        var bindables = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                bindableAttributeFullName,
                static (node, _) => node is ClassDeclarationSyntax or RecordDeclarationSyntax,
                static (context, cancel) => Locate(context, ExtractBindableTarget(context, cancel)))
            .Where(static _ => _ != null)
            .Select(static (target, _) => target!)
            .WithTrackingName(Stages.Bindables);

        context.RegisterSourceOutput(
            bindables,
            static (productionContext, target) => ProcessBindable(productionContext, target));
    }

    // Pairs the primitive snapshot with the syntax tree its attribute sits in, so the diagnostics
    // can be reported at a real source location. Only that one tree is carried, never the
    // compilation: the tree instance is replaced when its own file is edited and reused
    // otherwise, so an edit anywhere else still finds every stage cached.
    static LocatedTarget? Locate(GeneratorAttributeSyntaxContext context, TargetInfo? target)
    {
        if (target == null)
        {
            return null;
        }

        return new(target, context.Attributes.FirstOrDefault()?.ApplicationSyntaxReference?.SyntaxTree);
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

    static void ProcessBindable(SourceProductionContext context, LocatedTarget located)
    {
        var target = located.Target;
        var location = target.Location.ToLocation(located.Tree);

        if (ReportDeclarationError(context, target, location))
        {
            return;
        }

        // The template kind is unknown at compile time, so the shape rules that a docx
        // registration would need are enforced here — an invalid editable member is an authoring
        // mistake regardless of which template it later meets.
        ValidateShape(context, target, location, editableFields: true);
        ValidateEditableShape(context, target, location);
        ValidateExcelsiorShape(context, target, location);
        ValidateFormatMarkerTypes(context, target, location);

        EmitRegistration(context, target, GenerateBindableRegistration(target), "ParchmentBindable");
    }

    // The declarations nothing can be emitted for. Both attributes stop here, skipping validation
    // as well as registration: template tokens may still be valid, but reporting on a model that
    // will never be registered only buries the one diagnostic that matters.
    static bool ReportDeclarationError(SourceProductionContext context, TargetInfo target, Location location)
    {
        if (target.UnsupportedDeclaration != null)
        {
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.UnsupportedModelDeclaration,
                    location,
                    target.ModelDisplayName,
                    target.UnsupportedDeclaration));
            return true;
        }

        if (target.ExtractError != null)
        {
            // PARCH011: an enclosing type isn't partial. Emitting the registration helper into a
            // namespace-scope partial would land it in the wrong type.
            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.EnclosingTypeNotPartial,
                    location,
                    target.DeclaringName,
                    target.ExtractError));
            return true;
        }

        return false;
    }

    static string GenerateBindableRegistration(TargetInfo target)
    {
        var (fieldsBlock, registrationsBlock) = PrepareAccessors(target);
        // The type name is folded into the method name so a base and a derived model can both be
        // [ParchmentBindable] without the derived initializer hiding the base one (CS0108).
        return BuildPartialSource(
            target,
            fieldsBlock,
            $"InitializeParchmentBindable{target.DeclaringName}",
            registrationsBlock);
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
        var shape = ShapeBuilder.Build(typeSymbol, excelsiorTableType, editableFieldType, context.SemanticModel.Compilation, cancel);

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
            IsModuleAccessible(typeSymbol),
            new(enclosingResult.Chain.ToImmutableArray()),
            typeSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
            displayName,
            protection,
            EquatableLocation.From(rawLocation),
            shape,
            enclosingResult.Error,
            FindUnsupportedDeclaration(typeSymbol));
    }

    // The declarations no partial emitted from another file can register. A generic model has no
    // closed type to bind a template to, and a module initializer cannot sit inside a generic type
    // (CS8816), so neither it nor anything nested in one can register itself. A file-local type
    // cannot be extended at all: the generated partial is in another file, where the same name
    // declares a different type.
    static string? FindUnsupportedDeclaration(INamedTypeSymbol typeSymbol)
    {
        for (var current = typeSymbol; current != null; current = current.ContainingType)
        {
            var nested = !SymbolEqualityComparer.Default.Equals(current, typeSymbol);
            if (current.IsFileLocal)
            {
                if (nested)
                {
                    return $"it is nested inside the file-local type '{current.Name}', which a generated file cannot extend";
                }

                return "it is file-local, and a generated file cannot extend a file-local type";
            }

            if (current.TypeParameters.Length == 0)
            {
                continue;
            }

            if (nested)
            {
                return $"it is nested inside the generic type '{current.Name}', where a module initializer cannot be emitted";
            }

            return "it is generic, and a template binds to one closed type";
        }

        return null;
    }

    // Whether the declaration can be named from anywhere in the module. A module initializer has
    // to be (CS8814), along with every type it sits in.
    internal static bool IsModuleAccessible(INamedTypeSymbol typeSymbol) =>
        typeSymbol.DeclaredAccessibility is
            Accessibility.Public or
            Accessibility.Internal or
            Accessibility.ProtectedOrInternal;

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

            stack.Add(new(current.Name, GetTypeKindKeyword(current), IsModuleAccessible(current)));
        }

        // ContainingType walks innermost → outermost; flip so emission can write outermost first.
        stack.Reverse();
        return (stack, null);
    }

    internal static bool IsPartial(INamedTypeSymbol typeSymbol)
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

    internal static string GetTypeKindKeyword(INamedTypeSymbol type)
    {
        if (type.IsRecord)
        {
            if (type.TypeKind == TypeKind.Struct)
            {
                return "record struct";
            }

            return "record";
        }

        if (type.TypeKind == TypeKind.Struct)
        {
            return "struct";
        }

        if (type.TypeKind == TypeKind.Interface)
        {
            return "interface";
        }

        return "class";
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
            // Read once. A second read could land on the other side of a save and embed a template
            // other than the one validated.
            var bytes = ReadBytes(text.Path);
            var result = DocxArchiveReader.Read(bytes);
            return new(
                text.Path,
                new(result.Paragraphs.ToImmutableArray()),
                new(result.BodyParagraphs.ToImmutableArray()),
                new(result.BodyFallbacks.ToImmutableArray()),
                new(result.ParagraphContainers.ToImmutableArray()),
                new(result.BodyValueControlSpans.ToImmutableArray()),
                result.HasRemovePersonalInformation,
                Convert.ToBase64String(bytes),
                null);
        }
        catch (Exception exception)
        {
            return new(text.Path, EquatableArray<string>.Empty, EquatableArray<string>.Empty, EquatableArray<bool>.Empty, EquatableArray<int>.Empty, EquatableArray<int>.Empty, false, string.Empty, exception.Message);
        }
    }

    static DotxData ReadDotx(AdditionalText text)
    {
        try
        {
            // Opened the way a docx template is, so a file the runtime could not register as a
            // style source fails the build instead of the first render. Its paragraphs are not
            // kept: only the package's shape is checked here.
            var bytes = ReadBytes(text.Path);
            if (!DocxArchiveReader.Read(bytes).HasMainPart)
            {
                return new(text.Path, string.Empty, "it has no main document part (word/document.xml), so it is not a Word document or template");
            }

            return new(text.Path, Convert.ToBase64String(bytes), null);
        }
        catch (Exception exception)
        {
            return new(text.Path, string.Empty, exception.Message);
        }
    }

    // RS1035 steers analyzers to AdditionalText, but an AdditionalText has no binary read surface
    // — GetText() returns null for binary content — so raw bytes come off the stream directly.
    //
    // Shared for writing and for deleting. A word processor keeps the document it is editing open
    // with write access and replaces the file when it saves; asking for anything narrower gets the
    // open refused ("being used by another process") for a file that is perfectly readable, and
    // would in turn refuse a save that lands while the build is reading.
#pragma warning disable RS1035
    static byte[] ReadBytes(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
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
        LocatedTarget located,
        EquatableArray<DocxData> docs,
        EquatableArray<MarkdownData> markdowns,
        EquatableArray<DotxData> dotxes,
        string? projectDirectory,
        HashSet<string> reportedStyleDocs)
    {
        var target = located.Target;
        var location = target.Location.ToLocation(located.Tree);

        if (ReportDeclarationError(context, target, location))
        {
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
            ProcessMarkdown(context, target, location, markdownCandidates[0], dotxes, projectDirectory, reportedStyleDocs);
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

        var tokens = TokenScanner.Scan(matched.Paragraphs, matched.ParagraphContainers.AsImmutableArray());
        ValidateTokens(context, target, templatePath, tokens, location);
        ValidateBlockContainers(context, templatePath, tokens, location);
        ValidateShape(context, target, location, editableFields: true);

        // Editable fields are docx-only (like Excelsior / Format dispatch), so the shape rules
        // fire here rather than in Process — the runtime lockstep is EditableMap.Build throwing
        // at RegisterDocxTemplate regardless of whether a token references the member.
        ValidateEditableShape(context, target, location);
        ValidateExcelsiorShape(context, target, location);
        ValidateFormatMarkerTypes(context, target, location);
        if (HasEditableMembers(target.Shape))
        {
            // Token rules are body-scoped: the runtime dispatches editable fields only in the
            // document body, so a header/footer occurrence of the same member is a deliberate
            // read-only mirror, not a duplicate.
            var bodyTokens = TokenScanner.Scan(
                matched.BodyParagraphs,
                fallbacks: matched.BodyFallbacks.AsImmutableArray(),
                valueControlSpans: matched.BodyValueControlSpans.AsImmutableArray());
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
    static string? IgnoredStaticAttribute(MemberEntry member, bool editableFields)
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

        if (editableFields &&
            member.IsEditable)
        {
            return "EditableField";
        }

        return null;
    }

    // Its own pass rather than a branch of ValidateEditableShape: an Excelsior table renders in both
    // flows, so the markdown path runs this too. A static member is left to PARCH019, which already
    // says the attribute does nothing there.
    static void ValidateExcelsiorShape(
        SourceProductionContext context,
        TargetInfo target,
        Location location)
    {
        foreach (var type in target.Shape.Types)
        {
            foreach (var member in type.Members)
            {
                if (member is not {ExcelsiorNotTabular: true, IsStatic: false})
                {
                    continue;
                }

                context.ReportDiagnostic(
                    Diagnostic.Create(
                        Diagnostics.ExcelsiorNotTabular,
                        location,
                        target.ModelDisplayName,
                        $"{Display(type.TypeFullyQualifiedName)}.{member.Name}",
                        Display(member.TypeFullyQualifiedName)));
            }
        }
    }

    /// <summary>
    /// PARCH029: a format marker on a member that is not a string. Both flows read the marker, so
    /// this runs for a docx model, a markdown model and a bindable type alike. A static member is
    /// skipped: its marker is already reported as ignored.
    /// </summary>
    static void ValidateFormatMarkerTypes(
        SourceProductionContext context,
        TargetInfo target,
        Location location)
    {
        foreach (var type in target.Shape.Types)
        {
            foreach (var member in type.Members)
            {
                if (member.IsStatic ||
                    !member.FormatMarkerOnNonString)
                {
                    continue;
                }

                context.ReportDiagnostic(
                    Diagnostic.Create(
                        Diagnostics.FormatMarkerOnNonString,
                        location,
                        target.ModelDisplayName,
                        $"{Display(type.TypeFullyQualifiedName)}.{member.Name}",
                        Display(member.TypeFullyQualifiedName)));
            }
        }
    }

    /// <summary>
    /// The shape rules that hold whatever the model is bound to: a render attribute on a static
    /// member, an <c>[ExcelsiorTable]</c> Configure name that resolves to nothing, and format
    /// markers that contradict each other or the member's type.
    /// </summary>
    /// <remarks>
    /// All of them describe what the accessor emission does with the member, and that emission is
    /// the same pass for a docx template, a markdown template and a bindable model. The editable
    /// rules are separate (<see cref="ValidateEditableShape"/>) because a markdown template has no
    /// editable fields: <paramref name="editableFields"/> is false there, so a static member
    /// carrying only <c>[EditableField]</c> is as silent as an instance one.
    /// </remarks>
    static void ValidateShape(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        bool editableFields)
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
                    var ignored = IgnoredStaticAttribute(member, editableFields);
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

                if (member.ExcelsiorConfigureMissing)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.ExcelsiorConfigureMissing,
                            location,
                            target.ModelDisplayName,
                            memberDisplay,
                            member.ExcelsiorConfigure ?? "",
                            Display(type.TypeFullyQualifiedName)));
                }

                if (member.ExcelsiorConfigureInaccessibleOn != null)
                {
                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.ExcelsiorConfigureInaccessible,
                            location,
                            target.ModelDisplayName,
                            memberDisplay,
                            member.ExcelsiorConfigure ?? "",
                            Display(member.ExcelsiorConfigureInaccessibleOn)));
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
            }
        }
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
                if (member.IsStatic ||
                    !member.IsEditable)
                {
                    continue;
                }

                var memberDisplay = $"{Display(type.TypeFullyQualifiedName)}.{member.Name}";

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

        // An editable member below a struct has no setter that reaches the model. One report per
        // member, however many paths lead to it.
        var reported = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (declaringTypeFqn, member, valueTypeFqn) in AccessorEmission.EditablesThroughValueType(target.Shape))
        {
            var memberDisplay = $"{Display(declaringTypeFqn)}.{member.Name}";
            if (!reported.Add(memberDisplay))
            {
                continue;
            }

            context.ReportDiagnostic(
                Diagnostic.Create(
                    Diagnostics.EditableThroughValueType,
                    location,
                    target.ModelDisplayName,
                    memberDisplay,
                    Display(valueTypeFqn)));
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
        if (!member.EditableCollectionAssignable)
        {
            ReportCollectionInvalid(
                context,
                target,
                location,
                memberDisplay,
                $"its type '{Display(member.TypeFullyQualifiedName)}' cannot hold the rebuilt list — extraction builds a List<T> (a T[] for an array member), so declare the member as List<T>, T[], or an interface List<T> implements such as IList<T> or IReadOnlyList<T>");
        }

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
        if (element.IsValueType)
        {
            ReportCollectionInvalid(
                context,
                target,
                location,
                memberDisplay,
                $"its element type '{elementDisplay}' is a struct — extraction fills each element through a reference to it, and a struct is copied instead, so the element type must be a class");
            return;
        }

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
    /// <remarks>
    /// Lockstep with the runtime's <c>EditableTokenValidator</c>, and both follow where the runner
    /// holds an editable map: a root member outside every loop is a control, a member of an
    /// <c>[EditableField]</c> collection's item reached through the loop variable directly inside a
    /// top-level loop over that collection is a control, and every other occurrence renders
    /// read-only. PARCH016 / PARCH017 apply to the controls, PARCH018 to the rest.
    /// </remarks>
    static void ValidateEditableTokens(
        SourceProductionContext context,
        TargetInfo target,
        string templatePath,
        IReadOnlyList<Token> tokens,
        Location location)
    {
        var scope = new Dictionary<string, string>(StringComparer.Ordinal);
        var loopStack = new Stack<(string? Name, string? PriorBinding, bool IsEditableCollection, string? BoundCollection)>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // The item paths bound so far in the open repeating-section loop. Tags are item-relative,
        // so each top-level loop starts its own set.
        var itemSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // The collections some earlier loop already bound.
        var boundCollections = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var token in tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.ForOpen:
                    string? bound = null;
                    string? prior = null;
                    var isEditableCollectionLoop = false;
                    if (loopStack.Count == 0)
                    {
                        itemSeen.Clear();
                    }

                    if (token.LoopVariable != null &&
                        IsUntypedLoop(target.Shape, token, scope))
                    {
                        // Bound so the variable shadows a model member of the same name, as it
                        // does at render.
                        bound = token.LoopVariable;
                        prior = scope.GetValueOrDefault(bound);
                        scope[bound] = untypedLoopVariable;
                    }
                    else if (token is { LoopVariable: not null, References.Count: > 0 })
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

                    // Only a top-level loop renders as a repeating section — a nested one runs
                    // with an empty editable map at runtime and stays read-only — so that is the
                    // only place a modifier would be written back by extraction.
                    if (isEditableCollectionLoop &&
                        token.HasLoopModifiers &&
                        loopStack.Count == 0)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.EditableCollectionLoopModifier,
                                location,
                                templatePath,
                                token.Source,
                                string.Join('.', token.References[0])));
                    }

                    // Likewise only a top-level loop can bind the collection's controls.
                    string? boundCollection = null;
                    if (isEditableCollectionLoop &&
                        bound != null &&
                        loopStack.Count == 0)
                    {
                        boundCollection = string.Join('.', token.References[0]);
                    }

                    loopStack.Push((bound, prior, isEditableCollectionLoop, boundCollection));
                    break;

                case TokenKind.ForClose:
                    if (loopStack.Count > 0)
                    {
                        var (name, priorBinding, _, _) = loopStack.Pop();
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

                    // An editable collection is not a field: the runtime keeps it out of the
                    // scalar map, so its own token is an ordinary substitution.
                    if (!ShapeResolver.TryResolveMember(target.Shape, token.References[0], scope, out var member) ||
                        member is { IsEditable: false } or { IsStatic: true } or { EditableCollectionElementFqn: not null })
                    {
                        break;
                    }

                    var tokenSeen = seen;
                    if (loopStack.Count > 0)
                    {
                        // Directly inside a top-level loop over an editable collection, a token
                        // through the loop variable is a repeating-section control and answers to
                        // the same rules as a root one. Anything else inside a loop — a root
                        // member, a nested loop's token, a loop that is not a repeating section —
                        // renders read-only.
                        var frame = loopStack.Peek();
                        if (loopStack.Count > 1 ||
                            !frame.IsEditableCollection ||
                            !string.Equals(frame.Name, token.References[0][0], StringComparison.Ordinal))
                        {
                            context.ReportDiagnostic(
                                Diagnostic.Create(
                                    Diagnostics.EditableTokenInLoop,
                                    location,
                                    templatePath,
                                    token.Source));
                            break;
                        }

                        // The first control of a loop is what makes it a binding of the collection.
                        // Extraction reads one repeating section per collection, so a second loop
                        // binding any member is an edit that would be discarded. Lockstep with the
                        // runtime's EditableTokenValidator.
                        if (itemSeen.Count == 0 &&
                            frame.BoundCollection is { } boundPath &&
                            !boundCollections.Add(boundPath))
                        {
                            itemSeen.Add(string.Join('.', token.References[0]));
                            context.ReportDiagnostic(
                                Diagnostic.Create(
                                    Diagnostics.EditableTokenDuplicated,
                                    location,
                                    templatePath,
                                    token.Source));
                            break;
                        }

                        tokenSeen = itemSeen;
                    }

                    // What reaches here becomes a content control: a root member outside every
                    // loop, or an item member directly inside the top-level loop over an editable
                    // collection. Lockstep with EditableTokenValidator.RejectInsideValueControl.
                    if (token.InValueControl)
                    {
                        context.ReportDiagnostic(
                            Diagnostic.Create(
                                Diagnostics.EditableTokenInValueControl,
                                location,
                                templatePath,
                                token.Source));
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

                    // The fallback copy of a text box mirrors the choice and is skipped by
                    // extraction, so it is not a second use — as EditableTokenValidator.
                    if (!token.InFallback &&
                        !tokenSeen.Add(string.Join('.', token.References[0])))
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
    /// one is a property of where it sits in the source rather than of the template's AST — whether
    /// the token is alone on its line, and whether its expression is a bare path. Both are read off
    /// the source here, the same way <c>MarkdownExcelsiorTables</c> reads them at registration,
    /// through a copy of the same <c>MarkdownOutputTokens</c>.
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
        var loops = MarkdownLoopScope.Scan(markdown);
        foreach (var token in MarkdownOutputTokens.Scan(markdown, markdownParser))
        {
            // Root paths only, matching the map the render looks the table up in: a loop
            // variable's members are not reachable from the root and fall through to Fluid,
            // including when the loop variable happens to spell a root path. The text of a raw
            // or comment region is not a token at all.
            if (token.Path == null ||
                loops.IsLiteral(token.Start) ||
                loops.Shadows(token.Start, token.Path) ||
                !ShapeResolver.IsExcelsiorTableMember(target.Shape, token.Path.Split('.'), EmptyScope))
            {
                continue;
            }

            var source = markdown.Substring(token.Start, token.End - token.Start);
            if (!MarkdownOutputTokens.IsAlone(markdown, token, out _, out _))
            {
                context.ReportDiagnostic(
                    Diagnostic.Create(
                        Diagnostics.ExcelsiorTokenNotAlone,
                        location,
                        templatePath,
                        source));
                continue;
            }

            if (!token.IsPlain)
            {
                context.ReportDiagnostic(
                    Diagnostic.Create(
                        Diagnostics.ExcelsiorTokenNotPlainIdentifier,
                        location,
                        templatePath,
                        source));
            }
        }
    }

    static readonly Dictionary<string, string> EmptyScope = new();
    static void ProcessMarkdown(
        SourceProductionContext context,
        TargetInfo target,
        Location location,
        MarkdownData matched,
        EquatableArray<DotxData> dotxes,
        string? projectDirectory,
        HashSet<string> reportedStyleDocs)
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
                    BlockTagArguments.MarkdownParseFailure(matched.Text, error)));
            return;
        }

        MarkdownValidator.Validate(context, target, templatePath, location, template);
        ValidateFormatMarkerTypes(context, target, location);
        ValidateMarkdownExcelsiorTokens(context, target, templatePath, location, matched.Text);
        ValidateExcelsiorShape(context, target, location);

        // Editable fields are a docx-flow feature and are ignored here, so only the rules that
        // hold in every flow run: the format and Excelsior maps they guard are emitted for a
        // markdown template exactly as they are for a docx one.
        ValidateShape(context, target, location, editableFields: false);

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
            // Located at the style document, not at the model: the file is what needs fixing, and
            // every model that inherits it would otherwise repeat the same error.
            if (reportedStyleDocs.Add(style.Path))
            {
                context.ReportDiagnostic(
                    Diagnostic.Create(
                        Diagnostics.TemplateReadError,
                        Location.Create(style.Path, default, default),
                        DisplayPath(style.Path, projectDirectory),
                        style.ReadError));
            }

            return;
        }

        EmitRegistration(context, target, GenerateMarkdownRegistration(target, matched, style));
    }

    static readonly FluidParser markdownParser = new MarkdownFluidParser();

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

        var forwarders = AccessorEmission.EmitConfigureForwarders(target.Shape);
        if (forwarders != null)
        {
            context.AddSource(
                $"{hintPrefix}_{suffix}_Configure.g.cs",
                SourceText.From(forwarders, Encoding.UTF8));
        }
    }

    // Scope binding for a loop variable whose source is not statically known: a range, a path
    // rooted at another such variable, or one that passes an index the shape cannot see through.
    // The variable exists, but members reached through it are not checked. Mirrors `untyped` in
    // the runtime ReferenceValidator and the set in MarkdownValidator.
    const string untypedLoopVariable = ShapeResolver.Untyped;

    static bool IsUntypedLoop(ModelShape shape, Token token, IReadOnlyDictionary<string, string> scope)
    {
        if (token.LoopSourceIsRange)
        {
            return true;
        }

        if (token.References.Count == 0)
        {
            return false;
        }

        if (IsUntyped(token.References[0], scope))
        {
            return true;
        }

        // Reached through an index that says nothing about what it yields, a dictionary key for
        // one: the path is valid as far as it goes and the element is only known at render.
        return ShapeResolver.TryResolve(shape, token.References[0], scope, out var sourceFqn) &&
               sourceFqn == ShapeResolver.Untyped;
    }

    static bool IsUntyped(IReadOnlyList<string> reference, IReadOnlyDictionary<string, string> scope) =>
        scope.TryGetValue(reference[0], out var bound) &&
        bound == untypedLoopVariable;

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
                    ValidateFilterOrder(context, templatePath, location, token);
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

                    // `limit:` / `offset:` are evaluated before the loop variable is bound, so they
                    // resolve against the enclosing scope. `offset: continue` lands here too: the
                    // docx flow has no loop scope to resume from and the runtime rejects it.
                    if (token.LoopModifierReferences != null)
                    {
                        ValidateReferences(context, target, templatePath, location, token.LoopModifierReferences, scope, token.Source);
                    }

                    if (token.LoopVariable == null)
                    {
                        break;
                    }

                    if (IsUntypedLoop(target.Shape, token, scope))
                    {
                        // Nothing to resolve an element type from, which is an unknown, not a
                        // mistake. Members used as range bounds are still references like any
                        // other. Pushed like a typed loop so the closing tag pairs with this one.
                        ValidateReferences(context, target, templatePath, location, token.References, scope, token.Source);
                        var hadPriorUntyped = scope.TryGetValue(token.LoopVariable, out var priorUntyped);
                        scope[token.LoopVariable] = untypedLoopVariable;
                        loopStack.Push((token.LoopVariable, hadPriorUntyped ? priorUntyped : null));
                        break;
                    }

                    if (token.References.Count == 0)
                    {
                        break;
                    }

                    if (!ShapeResolver.TryResolve(target.Shape, token.References[0], scope, out var sourceFqn))
                    {
                        context.ReportDiagnostic(
                            Diagnostics.UnknownMember(target, templatePath, location, token.Source, token.References[0], scope));
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
                    var hint = string.Empty;
                    if (token.Problem != null)
                    {
                        hint = $". {token.Problem}";
                    }

                    context.ReportDiagnostic(
                        Diagnostic.Create(
                            Diagnostics.UnsupportedBlockTag,
                            location,
                            templatePath,
                            token.Source,
                            Diagnostics.DocxBlockTags,
                            hint));
                    break;
            }
        }
    }

    /// <summary>
    /// PARCH028: the tag paragraphs of one block — open, every elsif / else, close — must share a
    /// parent element. Lockstep with <c>ScopeTreeBuilder.RequireSameContainer</c> at runtime.
    /// </summary>
    /// <remarks>
    /// A pass of its own because <see cref="ValidateTokens"/> only keeps a stack for loops whose
    /// source resolved, and this rule is about where a tag sits rather than what it binds. Tags
    /// are paired the way the runtime pairs them, by document order. An unsupported tag leaves the
    /// nesting unknown — it may or may not open a block — so nothing is reported past PARCH003.
    /// </remarks>
    static void ValidateBlockContainers(
        SourceProductionContext context,
        string templatePath,
        IReadOnlyList<Token> tokens,
        Location location)
    {
        foreach (var token in tokens)
        {
            if (token.Kind == TokenKind.UnknownBlock)
            {
                return;
            }
        }

        var open = new Stack<(Token Opening, Token? Stray)>();
        foreach (var token in tokens)
        {
            switch (token.Kind)
            {
                case TokenKind.ForOpen:
                case TokenKind.IfOpen:
                    open.Push((token, null));
                    break;

                case TokenKind.ElsIf:
                case TokenKind.Else:
                    if (open.Count > 0 &&
                        open.Peek() is { Opening.Kind: TokenKind.IfOpen, Stray: null } block &&
                        token.Container != block.Opening.Container)
                    {
                        open.Pop();
                        open.Push((block.Opening, token));
                    }

                    break;

                case TokenKind.ForClose:
                    CloseBlock(context, templatePath, location, open, TokenKind.ForOpen, token);
                    break;

                case TokenKind.IfClose:
                    CloseBlock(context, templatePath, location, open, TokenKind.IfOpen, token);
                    break;
            }
        }
    }

    static void CloseBlock(
        SourceProductionContext context,
        string templatePath,
        Location location,
        Stack<(Token Opening, Token? Stray)> open,
        TokenKind opener,
        Token closing)
    {
        // A close with no matching open is malformed nesting, which is not this rule's to report.
        if (open.Count == 0 ||
            open.Peek().Opening.Kind != opener)
        {
            return;
        }

        var (opening, stray) = open.Pop();
        if (stray == null &&
            closing.Container != opening.Container)
        {
            stray = closing;
        }

        if (stray == null)
        {
            return;
        }

        context.ReportDiagnostic(
            Diagnostic.Create(
                Diagnostics.BlockTagsNotSiblings,
                location,
                templatePath,
                opening.Source,
                stray.Source));
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

    // PARCH037. The scanner has already found the pair — see TokenScanner.ParseSubstitution.
    static void ValidateFilterOrder(
        SourceProductionContext context,
        string templatePath,
        Location location,
        Token token)
    {
        if (token.TokenFilter == null ||
            token.FilterAfterTokenFilter == null)
        {
            return;
        }

        context.ReportDiagnostic(
            Diagnostic.Create(
                Diagnostics.FilterAfterTokenFilter,
                location,
                templatePath,
                token.Source,
                TokenFilterOrder.Message(token.TokenFilter, token.FilterAfterTokenFilter)));
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
            if (IsUntyped(reference, scope) ||
                ShapeResolver.TryResolve(target.Shape, reference, scope, out _))
            {
                continue;
            }

            context.ReportDiagnostic(
                Diagnostics.UnknownMember(target, templatePath, location, tokenSource, reference, scope));
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
        var statements =
            $$"""
              {{registrationsBlock}}  global::Parchment.Generated.GeneratedRegistration.RegisterDocxTemplate(
                  typeof({{target.ModelFullyQualifiedName}}),
                  global::System.Convert.FromBase64String("{{matched.ContentBase64}}"){{protection}});
              """;

        return BuildPartialSource(target, fieldsBlock, templateInitializer, statements);
    }

    static string GenerateMarkdownRegistration(TargetInfo target, MarkdownData matched, DotxData? style)
    {
        var (fieldsBlock, registrationsBlock) = PrepareAccessors(target);
        var markdown = SymbolDisplay.FormatLiteral(matched.Text, quote: true);
        var styleArgument = style == null
            ? ""
            : $",\n    global::System.Convert.FromBase64String(\"{style.ContentBase64}\")";
        var statements =
            $$"""
              {{registrationsBlock}}  global::Parchment.Generated.GeneratedRegistration.RegisterMarkdownTemplate(
                  typeof({{target.ModelFullyQualifiedName}}),
                  {{markdown}}{{styleArgument}});
              """;

        return BuildPartialSource(target, fieldsBlock, templateInitializer, statements);
    }

    const string templateInitializer = "InitializeParchmentTemplate";
    const string moduleInitializerAttribute = "[global::System.Runtime.CompilerServices.ModuleInitializer]";
    const string excludeFromCoverageAttribute = "[global::System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]";

    // Wraps the initializer in `partial {kind} {name} { ... }` declarations: namespace (if any),
    // then each enclosing type outermost-first, then the target itself. Indentation isn't strictly
    // necessary for correctness but keeps the generated source readable in obj/.../generated.
    static string BuildPartialSource(TargetInfo target, string fieldsBlock, string initializer, string statements)
    {
        // A private or protected nested model, or a model inside one, cannot carry the module
        // initializer itself (CS8814). The registration still has to live in the model, where the
        // model's own private members are in reach, so the attribute moves out to a forwarding
        // method in the nearest enclosing type the module can see — see AppendForwarder.
        var reachable = target.IsModuleAccessible;
        foreach (var enclosing in target.EnclosingTypes)
        {
            reachable &= enclosing.IsModuleAccessible;
        }

        var bodyBuilder = new StringBuilder(fieldsBlock);
        if (reachable)
        {
            bodyBuilder.Append(moduleInitializerAttribute).Append('\n');
        }

        bodyBuilder.Append($"internal static void {initializer}()\n{{\n");
        if (statements.Length > 0)
        {
            bodyBuilder.Append(statements.TrimEnd('\r', '\n')).Append('\n');
        }

        bodyBuilder.Append('}');
        var body = bodyBuilder.ToString();

        var builder = new StringBuilder(
            """
            // <auto-generated />
            #nullable enable

            """);

        // The accessors below name every bindable member, deprecated and experimental ones
        // included, and the consumer cannot edit this file to silence what that raises. Emitted
        // only when the model has such a member.
        if (target.Shape.SuppressedDiagnostics.Count > 0)
        {
            builder.Append("#pragma warning disable ");
            builder.AppendJoin(", ", target.Shape.SuppressedDiagnostics);
            builder.AppendLine();
        }

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
        // The model's full name is folded into the nested type's name so a derived model does not
        // hide its base's nested type (CS0108). The simple name alone is not enough: a model can
        // derive from one of the same name in another namespace. Internal rather than private
        // because a module initializer has to be reachable from the module, which a private nested
        // type is not.
        builder.Indent(depth + 1).AppendLine(excludeFromCoverageAttribute);
        builder.Indent(depth + 1).AppendLine($"internal static class {NestedTypeName(target)}");
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

        // How the initializer is named from inside the enclosing type being closed.
        var entry = $"{target.DeclaringName}.{NestedTypeName(target)}.{initializer}";
        var childAccessible = target.IsModuleAccessible;
        for (var i = depth - 1; i >= 0; i--)
        {
            if (!childAccessible)
            {
                entry = AppendForwarder(builder, target, i, initializer, entry);
            }

            builder.Indent(i).AppendLine("}");
            childAccessible = target.EnclosingTypes[i].IsModuleAccessible;
            entry = $"{target.EnclosingTypes[i].Name}.{entry}";
        }

        return builder.ToString();
    }

    // Emits, into the enclosing type at `index`, a method that calls the initializer one level
    // down, and returns how that method is named from inside the same enclosing type. One is
    // needed at every level whose nested type the module cannot see: a private type is only
    // nameable from inside the type that declares it, so each such level is stepped over from
    // directly above. The outermost of them is the one the module can reach, and carries the
    // attribute.
    static string AppendForwarder(StringBuilder builder, TargetInfo target, int index, string initializer, string entry)
    {
        var outermost = true;
        var name = new StringBuilder("ParchmentGenerated");
        for (var i = 0; i < target.EnclosingTypes.Count; i++)
        {
            // The path down to the model is folded into the name so two models under one
            // enclosing type do not declare the same forwarder.
            if (i > index)
            {
                name.Append(target.EnclosingTypes[i].Name).Append('_');
                continue;
            }

            // Index 0 is declared at namespace scope, which is never private.
            if (i > 0 &&
                !target.EnclosingTypes[i].IsModuleAccessible)
            {
                outermost = false;
            }
        }

        name.Append(target.DeclaringName).Append("Initializer");

        var depth = index + 1;
        builder.Indent(depth).AppendLine(excludeFromCoverageAttribute);
        builder.Indent(depth).AppendLine($"internal static class {name}");
        builder.Indent(depth).AppendLine("{");
        if (outermost)
        {
            builder.Indent(depth + 1).AppendLine(moduleInitializerAttribute);
        }

        builder.Indent(depth + 1).AppendLine($"internal static void {initializer}() =>");
        builder.Indent(depth + 2).AppendLine($"{entry}();");
        builder.Indent(depth).AppendLine("}");
        return $"{name}.{initializer}";
    }

    static string NestedTypeName(TargetInfo target)
    {
        const string globalPrefix = "global::";
        var name = target.ModelFullyQualifiedName;
        if (name.StartsWith(globalPrefix, StringComparison.Ordinal))
        {
            name = name[globalPrefix.Length..];
        }

        return "ParchmentGenerated_" + AccessorEmission.ToIdentifier(name);
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
