namespace Parchment;

public sealed class TemplateStore(ILogger<TemplateStore>? logger = null)
{
    ConcurrentDictionary<string, RegisteredTemplate> templates = new(StringComparer.Ordinal);
    ILogger logger = (ILogger?)logger ?? NullLogger.Instance;

    /// <summary>
    /// Policy for local-file image sources (<c>file://</c> URIs and filesystem paths) referenced
    /// from <c>&lt;img&gt;</c> tags or markdown <c>![alt](path)</c> images. Defaults to
    /// <see cref="OpenXmlHtml.ImagePolicy.AllowAll"/>.
    /// </summary>
    public ImagePolicy LocalImages { get; init; } = ImagePolicy.AllowAll();

    /// <summary>
    /// Policy for web image sources (<c>http://</c> and <c>https://</c> URIs). Defaults to
    /// <see cref="OpenXmlHtml.ImagePolicy.AllowAll"/>.
    /// </summary>
    public ImagePolicy WebImages { get; init; } = ImagePolicy.AllowAll();

    ImagePolicies Policies => new(LocalImages, WebImages);

    public void RegisterDocxTemplate<TModel>(string name, string path, ProtectionMode protection = ProtectionMode.WhenEditable)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var file = File.OpenRead(path);
        RegisterDocxTemplate<TModel>(name, file, protection);
    }

    public void RegisterDocxTemplate<TModel>(string name, Stream template, ProtectionMode protection = ProtectionMode.WhenEditable)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        GuardBindingModel<TModel>(name);
        SharedFluid.RegisterModel(typeof(TModel));

        var excelsiorMap = ExcelsiorTableMap.Build(typeof(TModel), name);
        var formatMap = FormatMap.Build(typeof(TModel), name);
        var stringListMap = StringListMap.Build(typeof(TModel));
        var editableMap = EditableMap.Build(typeof(TModel), name);

        using var stream = DocxCloner.ToWritableStream(template);
        IReadOnlyList<PartScopeTree> parts;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var bodyUri = doc.MainDocumentPart?.Uri.ToString();
            foreach (var (uri, root) in DocxCloner.EnumerateParts(doc))
            {
                var classifications = TokenScanner.Scan(root, name, uri);
                if (classifications.Count == 0)
                {
                    continue;
                }

                var tree = ScopeTreeBuilder.Build(classifications, name, uri);
                var validator = new ReferenceValidator(typeof(TModel), name, uri);
                validator.ValidateTree(tree);
                ExcelsiorTokenValidator.Validate(classifications, excelsiorMap, name, uri);
                FormatTokenValidator.Validate(classifications, formatMap, name, uri);
                EditableTokenValidator.Validate(classifications, editableMap, excelsiorMap, formatMap, name, uri, isBody: uri == bodyUri);
            }

            if (!editableMap.IsEmpty &&
                protection == ProtectionMode.WhenEditable)
            {
                SettingsProtection.Apply(doc.MainDocumentPart!);

                // Read-only protection locks the document-level numbering.xml and styles.xml outside
                // every editable range. Word then disables lists (numbering.xml) and the style gallery
                // (styles.xml) even inside a rich-text field the user may edit — unless the definitions
                // already exist. Seed both.
                if (editableMap.HasHtmlField)
                {
                    WordNumbering.EnsureListDefinitions(doc.MainDocumentPart!);
                    WordStyles.EnsureStyleDefinitions(doc.MainDocumentPart!);
                }
            }

            // Bake compatibilityMode=15 into the registration snapshot so every render inherits it
            // from the clone rather than re-stamping (a settings-part scan) on each render. It is
            // model-independent, so registration is the right place. See SettingsCompatibility.
            SettingsCompatibility.Apply(doc.MainDocumentPart!);

            doc.Save();

            parts = ExtractParts(doc, name);
        }

        var canonicalBytes = stream.ToArray();
        var registered = new RegisteredDocxTemplate(name, typeof(TModel), canonicalBytes, parts, excelsiorMap, formatMap, stringListMap, editableMap, Policies);
        templates[name] = registered;
        logger.LogInformation("Registered docx template {Name} for {ModelType}", name, typeof(TModel).Name);
    }

    public void RegisterMarkdownTemplate<TModel>(string name, string markdown, Stream? styleSource = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        GuardBindingModel<TModel>(name);
        SharedFluid.RegisterModel(typeof(TModel));

        if (!SharedFluid.Parser.TryParse(markdown, out var template, out var error))
        {
            throw new ParchmentRegistrationException(
                name,
                $"Failed to parse markdown as a liquid template: {error}");
        }

        var refs = IdentifierVisitor.Collect(template);
        foreach (var reference in refs)
        {
            if (reference.Segments.Count == 0)
            {
                continue;
            }

            // Skip loop variables — they get bound at render time via liquid scope and can't be
            // pre-validated without knowing the full scope tree.
            if (IsLikelyLoopVariable(reference.Root, markdown))
            {
                continue;
            }

            ModelValidator.Validate(typeof(TModel), reference, null, name, null, null);
        }

        byte[] bytes;
        if (styleSource is MemoryStream existingMs)
        {
            bytes = existingMs.ToArray();
        }
        else if (styleSource != null)
        {
            using var stream = new MemoryStream();
            styleSource.CopyTo(stream);
            bytes = stream.ToArray();
        }
        else
        {
            bytes = BlankDocxTemplate;
        }

        var registered = new RegisteredMarkdownTemplate(name, typeof(TModel), bytes, template, Policies);
        templates[name] = registered;
        logger.LogInformation("Registered markdown template {Name} for {ModelType}", name, typeof(TModel).Name);
    }

    static void GuardBindingModel<TModel>(string name)
    {
        var type = typeof(TModel);
        if (type.IsInterface)
        {
            throw new ParchmentRegistrationException(
                name,
                $"Model type '{type.Name}' is an interface. Parchment binds against a concrete type's properties via reflection — register against a class, record, or struct instead.");
        }
    }

    static byte[] BlankDocxTemplate { get; } = BuildBlankDocx();

    static byte[] BuildBlankDocx()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(new Paragraph()));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            styles.Append(
                new Style
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = "Normal",
                        Default = true
                    }
                    .AppendChild(
                        new StyleName
                        {
                            Val = "Normal"
                        }).Parent!);
            for (var i = 1; i <= 6; i++)
            {
                styles.Append(
                    new Style
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = $"Heading{i}"
                    }.AppendChild(
                        new StyleName
                        {
                            Val = $"Heading{i}"
                        }).Parent!);
            }

            styles.Append(
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "ListParagraph"
                }.AppendChild(
                    new StyleName
                    {
                        Val = "List Paragraph"
                    }).Parent!);
            styles.Append(
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Quote"
                }.AppendChild(
                    new StyleName
                    {
                        Val = "Quote"
                    }).Parent!);
            stylesPart.Styles = styles;
        }

        return stream.ToArray();
    }

    static bool IsLikelyLoopVariable(string identifier, string markdown)
    {
        // Hand-rolled scan for `{%\s*for\s+<identifier>\s+in\s+` so the markdown reference
        // validator can skip loop-bound names (they're bound at render time, not pre-validatable
        // against the model). Loop tag never spans paragraph boundaries in liquid, so a flat scan
        // of the markdown source is sufficient.
        var span = markdown.AsSpan();
        var idSpan = identifier.AsSpan();
        var index = 0;
        while (index < span.Length - 1)
        {
            var brace = span[index..].IndexOf('{');
            if (brace < 0)
            {
                return false;
            }

            index += brace;
            if (index + 1 >= span.Length ||
                span[index + 1] != '%')
            {
                index++;
                continue;
            }

            var cursor = index + 2;
            cursor = SkipWhitespace(span, cursor);
            if (!Matches(span, cursor, "for"))
            {
                index++;
                continue;
            }

            cursor += 3;
            var afterFor = SkipWhitespace(span, cursor);
            if (afterFor == cursor)
            {
                index++;
                continue;
            }

            cursor = afterFor;
            if (!Matches(span, cursor, idSpan))
            {
                index++;
                continue;
            }

            cursor += idSpan.Length;
            var afterId = SkipWhitespace(span, cursor);
            if (afterId == cursor)
            {
                index++;
                continue;
            }

            cursor = afterId;
            if (!Matches(span, cursor, "in"))
            {
                index++;
                continue;
            }

            cursor += 2;
            var afterIn = SkipWhitespace(span, cursor);
            if (afterIn == cursor)
            {
                index++;
                continue;
            }

            return true;
        }

        return false;
    }

    static int SkipWhitespace(CharSpan span, int from)
    {
        while (from < span.Length && char.IsWhiteSpace(span[from]))
        {
            from++;
        }

        return from;
    }

    static bool Matches(CharSpan span, int from, CharSpan literal) =>
        from + literal.Length <= span.Length &&
        span.Slice(from, literal.Length).SequenceEqual(literal);

    public Task Render(string name, object model, Stream output, Cancel cancel = default)
    {
        if (!templates.TryGetValue(name, out var template))
        {
            throw new ParchmentRenderException(name, "Template is not registered");
        }

        if (model == null)
        {
            throw new ParchmentRenderException(
                name,
                $"Model is null. Template is registered for {template.ModelType.Name}; pass a non-null instance.");
        }

        if (!template.ModelType.IsInstanceOfType(model))
        {
            throw new ParchmentRenderException(
                name,
                $"Model type mismatch: registered as {template.ModelType.Name} but received {model.GetType().Name}");
        }

        return template.Render(model, output, cancel);
    }

    public async Task RenderToFile(string name, object model, string path, Cancel cancel = default)
    {
        await using var file = File.Create(path);
        await Render(name, model, file, cancel).ConfigureAwait(false);
    }

    public static void AddFilter(string name, FilterDelegate filter) =>
        SharedFluid.Options.Filters.AddFilter(name, filter);

    static List<PartScopeTree> ExtractParts(WordprocessingDocument doc, string name)
    {
        var parts = new List<PartScopeTree>();
        foreach (var (uri, root) in DocxCloner.EnumerateParts(doc))
        {
            var classifications = new List<ParagraphClassification>();
            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                var anchorName = paragraph
                    .Elements<BookmarkStart>()
                    .FirstOrDefault(_ => _.Name?.Value?.StartsWith(Anchors.Prefix, StringComparison.Ordinal) == true)
                    ?.Name?.Value;
                if (anchorName == null)
                {
                    continue;
                }

                var reclassified = RestoreClassification(paragraph, anchorName);
                classifications.Add(reclassified);
            }

            if (classifications.Count == 0)
            {
                continue;
            }

            var tree = ScopeTreeBuilder.Build(classifications, name, uri);
            parts.Add(new(uri, tree));
        }

        return parts;
    }

    static ParagraphClassification RestoreClassification(
        Paragraph paragraph,
        string anchorName)
    {
        var text = ParagraphText.Build(paragraph);
        var innerText = text.InnerText;
        var sites = TokenScan.Scan(innerText);
        if (sites.Count == 0)
        {
            return new(paragraph, anchorName, ParagraphKind.Static, [], null);
        }

        var substitutions = new List<DocxTokenSite>();
        BlockMarker? block = null;
        foreach (var site in sites)
        {
            var source = innerText.Substring(site.Offset, site.Length);
            if (site.Kind == TokenSiteKind.Substitution)
            {
                if (SharedFluid.Parser.TryParse(source, out var template, out _))
                {
                    var refs = IdentifierVisitor.Collect(template);
                    substitutions.Add(new(site.Offset, site.Length, source, template, refs));
                }
            }
            else
            {
                block = ParseBlock(source);
            }
        }

        if (block != null)
        {
            return new(paragraph, anchorName, ParagraphKind.Block, [], block);
        }

        return new(paragraph, anchorName, ParagraphKind.Substitution, substitutions, null);
    }

    static BlockMarker? ParseBlock(string source)
    {
        if (!BlockTagParser.TryParse(source, out var tag, out var expression))
        {
            return null;
        }

        var expr = expression.IsEmpty ? null : expression.ToString();

        return tag switch
        {
            "for" => RebuildFor(source, expr),
            "endfor" => new(BlockTagKind.EndFor, source, null, null, null),
            "if" => RebuildIf(source, expr),
            "elsif" or "elseif" => RebuildElsif(source, expr),
            "else" => new(BlockTagKind.Else, source, null, null, null),
            "endif" => new(BlockTagKind.EndIf, source, null, null, null),
            _ => null
        };
    }

    static BlockMarker? RebuildFor(string source, string? expression)
    {
        if (expression == null)
        {
            return null;
        }

        var liquid = $"{{% for {expression} %}}{{% endfor %}}";
        if (!SharedFluid.Parser.TryParse(liquid, out var template, out _))
        {
            return null;
        }

        var forStatement = ((Fluid.Parser.FluidTemplate)template).Statements
            .OfType<ForStatement>()
            .FirstOrDefault();
        if (forStatement == null)
        {
            return null;
        }

        return new(BlockTagKind.For, source, null, forStatement.Identifier, forStatement.Source);
    }

    static BlockMarker? RebuildIf(string source, string? expression) =>
        RebuildConditional(BlockTagKind.If, source, expression);

    static BlockMarker? RebuildElsif(string source, string? expression) =>
        RebuildConditional(BlockTagKind.ElsIf, source, expression);

    static BlockMarker? RebuildConditional(BlockTagKind kind, string source, string? expression)
    {
        if (expression == null)
        {
            return null;
        }

        var liquid = $"{{% if {expression} %}}{{% endif %}}";
        if (!SharedFluid.Parser.TryParse(liquid, out var template, out _))
        {
            return null;
        }

        var ifStatement = ((Fluid.Parser.FluidTemplate)template).Statements
            .OfType<IfStatement>()
            .FirstOrDefault();
        if (ifStatement == null)
        {
            return null;
        }

        return new(kind, source, ifStatement.Condition, null, null);
    }
}
