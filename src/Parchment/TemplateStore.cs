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
        StaticRenderAttributes.Warn(typeof(TModel), name, logger);

        var excelsiorMap = ExcelsiorTableMap.Build(typeof(TModel), name);
        var formatMap = FormatMap.Build(typeof(TModel), name);
        var stringListMap = StringListMap.Build(typeof(TModel));
        var editableMap = EditableMap.Build(typeof(TModel), name);

        using var stream = DocxCloner.ToWritableStream(template);
        IReadOnlyList<PartScopeTree> parts;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            // Before anything reads the parts. ChangeDocumentType swaps the main part out, so any
            // scanning or rewriting done first would be thrown away with the part it was done on —
            // which is how a .dotx form template silently kept its FORMTEXT fields.
            EnsureDocumentType(doc);

            var bodyUri = doc.MainDocumentPart?.Uri.ToString();
            foreach (var (uri, root) in DocxCloner.EnumerateParts(doc))
            {
                // A template authored as a Word form carries its placeholders as FORMTEXT fields
                // rather than tokens. Rewrite them first so the rest of registration sees a normal
                // docx template — the mutation is baked into the registration snapshot below.
                FormFields.ToTokens(root);

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
        StaticRenderAttributes.Warn(typeof(TModel), name, logger);

        if (!SharedFluid.Parser.TryParse(markdown, out var template, out var error))
        {
            throw new ParchmentRegistrationException(
                name,
                $"Failed to parse markdown as a liquid template: {error}");
        }

        MarkdownReferenceValidator.Validate(template, typeof(TModel), name);

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

        bytes = NormalizeStyleSource(bytes);
        var (styleSourceBytes, parts) = ScanNonBodyParts<TModel>(bytes, name);

        var registered = new RegisteredMarkdownTemplate(name, typeof(TModel), styleSourceBytes, parts, template, Policies);
        templates[name] = registered;
        logger.LogInformation("Registered markdown template {Name} for {ModelType}", name, typeof(TModel).Name);
    }

    /// <summary>
    /// Binds the style source's non-body parts — headers, footers, notes — the way the docx flow
    /// binds every part.
    /// </summary>
    /// <remarks>
    /// The markdown replaces the body and nothing else, so every other part arrives from the style
    /// source exactly as authored. A style source is normally a real <c>.dotx</c>, and a header
    /// carrying a date or a reference is ordinary, so those tokens have to bind too — otherwise the
    /// same token binds in one flow and renders as literal text in the other.
    ///
    /// The scan mutates the package: anchors are baked in here so the render can find each scope
    /// again, which is why the rewritten bytes are returned rather than the ones passed in.
    /// </remarks>
    static (byte[] Bytes, IReadOnlyList<PartScopeTree> Parts) ScanNonBodyParts<TModel>(byte[] bytes, string name)
    {
        using var stream = DocxCloner.ToWritableStream(bytes);
        List<PartScopeTree> parts;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var bodyUri = doc.MainDocumentPart?.Uri.ToString();
            var found = false;
            foreach (var (uri, root) in DocxCloner.EnumerateParts(doc))
            {
                if (uri == bodyUri)
                {
                    continue;
                }

                var classifications = TokenScanner.Scan(root, name, uri);
                if (classifications.Count == 0)
                {
                    continue;
                }

                var tree = ScopeTreeBuilder.Build(classifications, name, uri);
                new ReferenceValidator(typeof(TModel), name, uri).ValidateTree(tree);
                found = true;
            }

            if (!found)
            {
                return (bytes, []);
            }

            doc.Save();
            parts = ExtractParts(doc, name);
            // ExtractParts walks every part. The body was never scanned, so whatever it found there
            // belongs to the markdown that replaces it, not to this pass.
            parts.RemoveAll(_ => _.PartUri == bodyUri);
        }

        return (stream.ToArray(), parts);
    }

    /// <summary>
    /// Retypes a template-typed package (<c>.dotx</c>/<c>.dotm</c>) as a document.
    /// </summary>
    /// <remarks>
    /// Both flows clone the supplied package and the clone becomes the rendered output, so a
    /// template source would otherwise produce a template-typed output — which Word opens as a new
    /// unsaved document based on it, rather than as the document itself. The type is
    /// model-independent, so this belongs at registration rather than on every render.
    /// </remarks>
    static void EnsureDocumentType(WordprocessingDocument doc)
    {
        if (doc.DocumentType != WordprocessingDocumentType.Document)
        {
            doc.ChangeDocumentType(WordprocessingDocumentType.Document);
        }
    }

    static byte[] NormalizeStyleSource(byte[] bytes)
    {
        using var stream = DocxCloner.ToWritableStream(bytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            if (doc.DocumentType == WordprocessingDocumentType.Document)
            {
                return bytes;
            }

            EnsureDocumentType(doc);
            doc.Save();
        }

        return stream.ToArray();
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

    public Task Render(string name, object model, Stream output, Cancel cancel = default) =>
        Render(name, model, output, null, cancel);

    /// <summary>
    /// Renders and stamps the document's properties. Only the values set on
    /// <paramref name="properties"/> are written; each part is merged, so properties the template
    /// carries of its own survive.
    /// </summary>
    public Task Render(string name, object model, Stream output, WordDocumentProperties? properties, Cancel cancel = default)
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

        return template.Render(model, output, properties, cancel);
    }

    public Task RenderToFile(string name, object model, string path, Cancel cancel = default) =>
        RenderToFile(name, model, path, null, cancel);

    /// <inheritdoc cref="Render(string, object, Stream, WordDocumentProperties, Cancel)"/>
    public async Task RenderToFile(string name, object model, string path, WordDocumentProperties? properties, Cancel cancel = default)
    {
        await using var file = File.Create(path);
        await Render(name, model, file, properties, cancel).ConfigureAwait(false);
    }

    public static void AddFilter(string name, FilterDelegate filter) =>
        SharedFluid.AddFilter(name, filter);

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
