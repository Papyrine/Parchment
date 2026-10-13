namespace Parchment;

public sealed class TemplateStore(ILogger<TemplateStore>? logger = null)
{
    ConcurrentDictionary<Type, RegisteredTemplate> templates = new();
    ILogger logger = (ILogger?)logger ?? NullLogger.Instance;
    System.Threading.Lock materializeLock = new();

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

    /// <summary>
    /// Resolves the page each bookmark falls on, filling in a table of contents and any page
    /// references as the document is built. Defaults to null, which leaves both for Word to compute
    /// when the reader accepts its prompt on open.
    /// </summary>
    public IPageNumberResolver? PageNumbers { get; init; }

    ImagePolicies Policies => new(LocalImages, WebImages);

    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public void RegisterDocxTemplate<TModel>(string path, ProtectionMode protection = ProtectionMode.WhenEditable)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var file = File.OpenRead(path);
        RegisterDocxTemplate<TModel>(file, protection);
    }

    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public void RegisterDocxTemplate<TModel>(Stream template, ProtectionMode protection = ProtectionMode.WhenEditable)
    {
        ArgumentNullException.ThrowIfNull(template);
        GuardReadable(template, typeof(TModel).Name, "template");
        RegisterDocxTemplate(typeof(TModel), DocxCloner.ToWritableStream(template), protection);
    }

    /// <summary>
    /// Registers <paramref name="clone"/> as the template for <paramref name="modelType"/>.
    /// </summary>
    /// <remarks>
    /// The caller supplies the writable clone and this method owns it: registration mutates the
    /// package in place — retyping it, rewriting form fields, baking in anchors — and keeps the
    /// result as the snapshot every render starts from. Taking the clone rather than making one
    /// spares the source-generator path a round-trip, since a definition already holds bytes.
    /// </remarks>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    void RegisterDocxTemplate(Type modelType, MemoryStream clone, ProtectionMode protection)
    {
        var name = modelType.Name;
        GuardBindingModel(modelType, name);
        SharedFluid.EnsureModelRegistered(modelType, name);

        var excelsiorMap = ExcelsiorTableMap.Build(modelType);
        var formatMap = FormatMap.Build(modelType);
        var stringListMap = StringListMap.Build(modelType);
        var editableMap = EditableMap.Build(modelType);

        using var stream = clone;
        var parts = new List<PartScopeTree>();
        var maxBookmarkId = 0;
        var templateIds = new TemplateIds();
        var noteReferences = NoteReferences.None;
        bool macros;
        var maxEditableId = 0;
        using (var doc = OpenTemplate(stream, name, "template"))
        {
            // Before anything reads the parts. ChangeDocumentType swaps the main part out, so any
            // scanning or rewriting done first would be thrown away with the part it was done on —
            // which is how a .dotx form template silently kept its FORMTEXT fields.
            EnsureDocumentType(doc, out macros);

            var bodyUri = doc.MainDocumentPart?.Uri.ToString();
            foreach (var (uri, _, root) in DocxCloner.EnumerateParts(doc))
            {
                if (uri == bodyUri)
                {
                    // Which notes the body refers to before any render repeats or removes a
                    // reference - see NoteReferences.
                    noteReferences = NoteReferences.Read(doc.MainDocumentPart!, root);
                }

                // A template authored as a Word form carries its placeholders as FORMTEXT fields
                // rather than tokens. Rewrite them first so the rest of registration sees a normal
                // docx template — the mutation is baked into the registration snapshot below.
                FormFields.ToTokens(root);

                // A render reads only the parts that carry tokens, so a cell the template itself
                // left without a closing paragraph is closed here, in every part, and not left to
                // the render's own pass, which never sees a token-free part.
                TableCells.EnsureFinalParagraph(root);

                // An anchor the template arrived with names nothing in this registration, and in a
                // part without tokens no render would ever reach it to take it out.
                Anchors.StripArrived(root);

                var classifications = TokenScanner.Scan(root, name, uri);

                // After the scan, so the anchors it added are counted, and across every part: a
                // render numbers what it produces from above this - see BookmarkIds.
                maxBookmarkId = Math.Max(maxBookmarkId, Anchors.MaxBookmarkId(root));

                // Likewise the content controls and editable ranges the template already holds, in
                // every part: an editable field is numbered from above them - see EditableState.
                if (!editableMap.IsEmpty)
                {
                    maxEditableId = Math.Max(maxEditableId, EditableState.MaxId(root));
                }

                // Every part for the same reason: what a loop copies is numbered clear of these.
                templateIds.Read(root);
                if (classifications.Count == 0)
                {
                    continue;
                }

                var tree = ScopeTreeBuilder.Build(classifications, name, uri);
                var validator = new ReferenceValidator(modelType, name, uri);
                validator.ValidateTree(tree);
                ExcelsiorTokenValidator.Validate(classifications, excelsiorMap, name, uri);
                FormatTokenValidator.Validate(classifications, formatMap, name, uri);
                EditableTokenValidator.Validate(classifications, editableMap, excelsiorMap, formatMap, name, uri, isBody: uri == bodyUri);

                // The tree that was just validated is the one the render walks. It names the
                // anchors this scan wrote into the package saved below, so nothing has to be read
                // back out of the package to build it.
                parts.Add(new(uri, tree));
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
        }

        if (macros)
        {
            MacroParts.RemoveContentTypeDefaults(stream);
        }

        var canonicalBytes = stream.ToArray();
        var registered = new RegisteredDocxTemplate(name, modelType, canonicalBytes, parts, excelsiorMap, formatMap, stringListMap, editableMap, Policies, maxBookmarkId, noteReferences, templateIds, maxEditableId);
        templates[modelType] = registered;
        logger.LogInformation("Registered docx template for {ModelType}", modelType.Name);
    }

    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public void RegisterMarkdownTemplate<TModel>(string markdown, Stream? styleSource = null)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        byte[]? bytes = null;
        if (styleSource != null)
        {
            GuardReadable(styleSource, typeof(TModel).Name, "style source");
            bytes = DocxCloner.ToBytes(styleSource);
        }

        RegisterMarkdownTemplate(typeof(TModel), markdown, bytes);
    }

    /// <summary>
    /// Registers <paramref name="markdown"/> as the template for <paramref name="model"/>, styled by
    /// <paramref name="styleSource"/> or by <see cref="BlankDocxTemplate"/> when none is supplied.
    /// </summary>
    /// <remarks>
    /// The style source arrives as an array rather than a stream so that
    /// <see cref="PrepareStyleSource"/> can hand it straight back when it has nothing to rewrite —
    /// which is the common case — and <see cref="BakeStyleSource"/> then makes the one copy the
    /// registered template keeps. The caller's array is never written to, which matters most for
    /// the source-generator path, where it belongs to a definition that outlives every store.
    /// </remarks>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    void RegisterMarkdownTemplate(Type model, string markdown, byte[]? styleSource)
    {
        var name = model.Name;
        GuardBindingModel(model, name);
        SharedFluid.EnsureModelRegistered(model, name);

        // Excelsior tokens become markers before the source is parsed, so the cached template
        // carries the markers and a render never has to re-scan the markdown.
        var excelsiorTables = ExcelsiorTableMap.Build(model);
        var (rewritten, tablePlaceholders) = MarkdownExcelsiorTables.Rewrite(markdown, excelsiorTables, name);

        // [Html] / [Markdown] members become filtered tokens for the same reason, and after the
        // Excelsior pass so a marked line is already a marker and cannot match as a plain token.
        rewritten = MarkdownFormats.Rewrite(rewritten, FormatMap.Build(model));

        if (!SharedFluid.Parser.TryParse(rewritten, out var template, out var error))
        {
            throw new ParchmentRegistrationException(
                name,
                BlockTagArguments.MarkdownParseFailure(rewritten, error));
        }

        MarkdownReferenceValidator.Validate(template, model, name);
        template = MarkdownNumberText.Rewrite(template);

        var (prepared, parts, maxBookmarkId, templateIds) = PrepareStyleSource(model, styleSource ?? BlankDocxTemplate, name);
        var (styleSourceBytes, compatibilityBaked) = BakeStyleSource(prepared);

        var registered = new RegisteredMarkdownTemplate(
            name,
            model,
            styleSourceBytes,
            parts,
            // After validation, which reads the statements as the author wrote them.
            MarkdownBlankOutputs.Instrument(template),
            Policies,
            PageNumbers,
            excelsiorTables,
            tablePlaceholders,
            maxBookmarkId,
            templateIds,
            compatibilityBaked);
        templates[model] = registered;
        logger.LogInformation("Registered markdown template for {ModelType}", model.Name);
    }

    /// <summary>
    /// Retypes the style source as a document and binds its non-body parts — headers, footers,
    /// notes — the way the docx flow binds every part.
    /// </summary>
    /// <remarks>
    /// The markdown replaces the body and nothing else, so every other part arrives from the style
    /// source exactly as authored. A style source is normally a real <c>.dotx</c>, and a header
    /// carrying a date or a reference is ordinary, so those tokens have to bind too — otherwise the
    /// same token binds in one flow and renders as literal text in the other.
    ///
    /// Both jobs share one open, as they do in the docx flow: they were a pass each, so the package
    /// was cloned and fully parsed twice to do work a single visit covers. Retyping still comes
    /// first for the reason <see cref="RegisterDocxTemplate(Type, MemoryStream, ProtectionMode)"/>
    /// gives — <c>ChangeDocumentType</c> swaps the main part out, so a scan done before it would be
    /// discarded along with the part it ran on.
    ///
    /// The scan mutates the package: anchors are baked in here so the render can find each scope
    /// again, which is why the rewritten bytes are returned rather than the ones passed in. When
    /// nothing changed — an already-document style source with no tokens outside its body and no
    /// anchor left in it from an earlier registration, which includes
    /// <see cref="BlankDocxTemplate"/> — the input is handed back untouched and nothing is
    /// serialized.
    /// </remarks>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    static (byte[] Bytes, IReadOnlyList<PartScopeTree> Parts, int MaxBookmarkId, TemplateIds Ids) PrepareStyleSource(Type model, byte[] bytes, string name)
    {
        using var stream = DocxCloner.ToWritableStream(bytes);
        var parts = new List<PartScopeTree>();
        var maxBookmarkId = 0;
        var templateIds = new TemplateIds();
        bool macros;
        using (var doc = OpenTemplate(stream, name, "style source"))
        {
            var retyped = EnsureDocumentType(doc, out macros);

            var bodyUri = doc.MainDocumentPart?.Uri.ToString();
            var found = false;
            var stripped = false;
            foreach (var (uri, _, root) in DocxCloner.EnumerateParts(doc))
            {
                if (uri == bodyUri)
                {
                    continue;
                }

                // An anchor the style source arrived with names nothing in this registration. A
                // render reads only the parts listed below, so it is taken out here, into the bytes
                // every render clones, and is a reason to rewrite the package even with no token.
                stripped |= Anchors.StripArrived(root);

                var classifications = TokenScanner.Scan(root, name, uri);

                // The body is skipped above because the markdown replaces it, so its bookmarks never
                // reach the output. Every other part's do, anchors included - see BookmarkIds.
                maxBookmarkId = Math.Max(maxBookmarkId, Anchors.MaxBookmarkId(root));

                // The same parts, for the ids a header or footer loop gives its copies.
                templateIds.Read(root);
                if (classifications.Count == 0)
                {
                    continue;
                }

                var tree = ScopeTreeBuilder.Build(classifications, name, uri);
                new ReferenceValidator(model, name, uri).ValidateTree(tree);
                parts.Add(new(uri, tree));
                found = true;
            }

            if (!found &&
                !retyped &&
                !stripped)
            {
                if (!retyped)
                {
                    return (bytes, [], maxBookmarkId, templateIds);
                }

                // Retyping alone leaves no token to bind, so no part is listed, as when nothing
                // was rewritten at all.
                parts.Clear();
            }

            doc.Save();
        }

        if (macros)
        {
            MacroParts.RemoveContentTypeDefaults(stream);
        }

        return (stream.ToArray(), parts, maxBookmarkId, templateIds);
    }

    /// <summary>
    /// Empties the style source's body and stamps the compatibility setting, so a render starts
    /// from a package already in the state every render used to rebuild.
    /// </summary>
    /// <remarks>
    /// The markdown replaces the body, and a style source authored in Word can carry a whole sample
    /// document there: parsing that on every render only to discard it was the larger part of a
    /// small render. Neither step depends on the model, so both belong at registration. What is
    /// left in the body is the one <c>w:sectPr</c> a render preserves, as a direct child, which is
    /// where a render looks for it. Parts the discarded body referenced stay in the package, as
    /// they always have.
    /// </remarks>
    internal static (byte[] Bytes, bool CompatibilityBaked) BakeStyleSource(byte[] bytes)
    {
        bool compatibilityBaked;
        using var stream = DocxCloner.ToWritableStream(bytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart;
            if (mainPart?.Document?.Body is not { } body)
            {
                // Registration never gets here: OpenTemplate has already rejected such a source.
                return (bytes, false);
            }

            var sectPr = body.Elements<SectionProperties>().LastOrDefault()
                         ?? body.Descendants<SectionProperties>().LastOrDefault();
            sectPr?.Remove();
            body.RemoveAllChildren();
            if (sectPr != null)
            {
                body.AppendChild(sectPr);
            }

            // Only into a settings part the style source already has. One created here would sit
            // ahead of whatever a render adds - numbering, images - in the package and in the
            // relationships, where a render has always added it after them, and the output would
            // no longer be the same bytes. Creating it is cheap; parsing an authored one is not.
            compatibilityBaked = mainPart.DocumentSettingsPart != null;
            if (compatibilityBaked)
            {
                SettingsCompatibility.Apply(mainPart);
            }

            doc.Save();
        }

        return (stream.ToArray(), compatibilityBaked);
    }

    /// <summary>
    /// Makes the package a plain, macro-free document: retypes a template-typed or macro-enabled
    /// one (<c>.dotx</c>/<c>.dotm</c>/<c>.docm</c>) and removes the VBA project it carries.
    /// </summary>
    /// <remarks>
    /// Both flows clone the supplied package and the clone becomes the rendered output, so a
    /// template source would otherwise produce a template-typed output — which Word opens as a new
    /// unsaved document based on it, rather than as the document itself. The type is
    /// model-independent, so this belongs at registration rather than on every render.
    ///
    /// The macros go with the type, also from a source that arrived typed as a document and
    /// carrying them anyway - see <see cref="MacroParts"/>. <paramref name="macros"/> tells the
    /// caller to finish that with <see cref="MacroParts.RemoveContentTypeDefaults"/> once the
    /// package is saved and closed, which is the only point the content types can be rewritten.
    /// </remarks>
    /// <returns>Whether the package was changed.</returns>
    static bool EnsureDocumentType(WordprocessingDocument doc, out bool macros)
    {
        var type = doc.DocumentType;
        macros = type is WordprocessingDocumentType.MacroEnabledDocument or WordprocessingDocumentType.MacroEnabledTemplate;
        var retyped = type != WordprocessingDocumentType.Document;
        if (retyped)
        {
            doc.ChangeDocumentType(WordprocessingDocumentType.Document);
        }

        // After the retype, which swaps the main part out.
        if (doc.MainDocumentPart is { } mainPart &&
            MacroParts.Remove(mainPart))
        {
            macros = true;
        }

        return retyped || macros;
    }

    /// <summary>
    /// Rejects a stream nothing can be read from, before it is copied.
    /// </summary>
    /// <remarks>
    /// A closed stream reports that it can neither read nor seek, so without this it surfaced as
    /// an <see cref="ObjectDisposedException"/> or <see cref="NotSupportedException"/> from the
    /// copy, naming neither the template nor which argument was at fault.
    /// </remarks>
    static void GuardReadable(Stream stream, string name, string subject)
    {
        if (!stream.CanRead)
        {
            throw new ParchmentRegistrationException(
                name,
                $"The {subject} stream cannot be read. It is closed, disposed or write-only.");
        }
    }

    /// <summary>
    /// Opens <paramref name="stream"/> as a Word package, or says why it is not one.
    /// </summary>
    /// <remarks>
    /// Every registration goes through here — both flows, and the definitions the source generator
    /// stores — so an unusable input fails at registration, as a
    /// <see cref="ParchmentRegistrationException"/> naming the template, and never as a
    /// <see cref="NullReferenceException"/> or at the first render. The SDK opens an empty stream,
    /// and a zip that is not a package, as a package with no parts, so opening without error
    /// proves nothing on its own.
    ///
    /// The body is read here for both flows. The markdown flow replaces it, but a render still
    /// needs one to fill, and enumerating the parts loads it regardless.
    /// </remarks>
    static WordprocessingDocument OpenTemplate(MemoryStream stream, string name, string subject)
    {
        if (stream.Length == 0)
        {
            throw new ParchmentRegistrationException(
                name,
                $"The {subject} is empty (0 bytes). A Word document (.docx or .dotx) is required.");
        }

        WordprocessingDocument doc;
        try
        {
            doc = WordprocessingDocument.Open(stream, true);
        }
        catch (Exception exception)
        {
            throw new ParchmentRegistrationException(
                name,
                $"The {subject} is not a Word document (.docx or .dotx). {DescribeUnopenable(stream)}: {exception.Message}",
                inner: exception);
        }

        try
        {
            GuardWordPackage(doc, name, subject);
        }
        catch
        {
            doc.Dispose();
            throw;
        }

        return doc;
    }

    static string DescribeUnopenable(MemoryStream stream)
    {
        // The header of an OLE compound file, which is what a legacy .doc is, and what Word wraps
        // a password-protected .docx in.
        ReadOnlySpan<byte> compoundFile = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        if (stream.TryGetBuffer(out var buffer) &&
            buffer.AsSpan().StartsWith(compoundFile))
        {
            return "It is a legacy .doc or a password-protected document, neither of which is an Open XML package";
        }

        return "It could not be opened as an Open XML package";
    }

    static void GuardWordPackage(WordprocessingDocument doc, string name, string subject)
    {
        if (doc.MainDocumentPart is not { } mainPart)
        {
            throw new ParchmentRegistrationException(
                name,
                $"The {subject} is not a Word document (.docx or .dotx): it has no main document part.");
        }

        var partUri = mainPart.Uri.ToString();
        Document? document;
        try
        {
            document = mainPart.Document;
        }
        catch (Exception exception)
        {
            throw new ParchmentRegistrationException(
                name,
                $"The main document part of the {subject} could not be read as a Word document. {DescribeUnreadable(mainPart)}: {exception.Message}",
                partUri,
                inner: exception);
        }

        if (document?.Body == null)
        {
            throw new ParchmentRegistrationException(
                name,
                $"The {subject} has no document body.",
                partUri);
        }
    }

    static string DescribeUnreadable(MainDocumentPart mainPart)
    {
        var contentType = mainPart.ContentType;
        if (contentType.Contains("spreadsheetml", StringComparison.OrdinalIgnoreCase))
        {
            return "It is a spreadsheet (.xlsx)";
        }

        if (contentType.Contains("presentationml", StringComparison.OrdinalIgnoreCase))
        {
            return "It is a presentation (.pptx)";
        }

        return "Its XML is malformed or is not WordprocessingML";
    }

    static void GuardBindingModel(Type type, string name)
    {
        if (type.IsInterface)
        {
            throw new ParchmentRegistrationException(
                name,
                $"Model type '{type.Name}' is an interface. Parchment binds against a concrete type's properties via reflection — register against a class, record, or struct instead.");
        }
    }

    static byte[] BlankDocxTemplate { get; } = BuildBlankDocx();

    internal static byte[] BuildBlankDocx()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            // Explicit ids: without them the SDK mints random ones, and since this is built once per
            // process, default markdown output would differ between processes. The package is new,
            // so nothing can already hold either id.
            var mainPart = doc.AddMainDocumentPart();
            doc.ChangeIdOfPart(mainPart, "rId1");
            mainPart.Document = new(new Body(new Paragraph()));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>("rIdStyles");
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
                styles.Append(BuildHeadingStyle(i));
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
            styles.Append(BuildQuoteStyle());
            stylesPart.Styles = styles;
        }

        return stream.ToArray();
    }

    /// <summary>
    /// Builds one of the blank template's heading styles.
    /// </summary>
    /// <remarks>
    /// A style with only a name renders as body text, so markdown registered without a style source
    /// produced a document whose headings were indistinguishable from its paragraphs. The formatting
    /// here is deliberately plain — a real style source overrides all of it — but it has to be
    /// present for the built-in fallback to read as a document rather than as a wall of text.
    /// </remarks>
    static Style BuildHeadingStyle(int level)
    {
        // 16pt down to 11pt, in half-points.
        var size = level switch
        {
            1 => 32,
            2 => 26,
            3 => 24,
            4 => 24,
            5 => 22,
            _ => 22
        };

        return new()
        {
            Type = StyleValues.Paragraph,
            StyleId = $"Heading{level}",
            StyleName = new()
            {
                Val = $"Heading{level}"
            },
            BasedOn = new()
            {
                Val = "Normal"
            },
            NextParagraphStyle = new()
            {
                Val = "Normal"
            },
            PrimaryStyle = new(),
            StyleParagraphProperties = new(
                new KeepNext(),
                new SpacingBetweenLines
                {
                    Before = "240",
                    After = "120"
                }),
            StyleRunProperties = new(
                new Bold(),
                new Color
                {
                    Val = "2F5496"
                },
                new FontSize
                {
                    Val = size.ToString()
                },
                new FontSizeComplexScript
                {
                    Val = size.ToString()
                })
        };
    }

    /// <summary>
    /// Builds the blank template's blockquote style, on the same reasoning as
    /// <see cref="BuildHeadingStyle"/>.
    /// </summary>
    static Style BuildQuoteStyle() =>
        new()
        {
            Type = StyleValues.Paragraph,
            StyleId = "Quote",
            StyleName = new()
            {
                Val = "Quote"
            },
            BasedOn = new()
            {
                Val = "Normal"
            },
            NextParagraphStyle = new()
            {
                Val = "Normal"
            },
            PrimaryStyle = new(),
            StyleParagraphProperties = new(
                new SpacingBetweenLines
                {
                    Before = "200",
                    After = "200"
                },
                // Half an inch of air on both sides, in twentieths of a point.
                new Indentation
                {
                    Left = "720",
                    Right = "720"
                }),
            StyleRunProperties = new(
                new Italic(),
                new Color
                {
                    Val = "404040"
                })
        };

    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public Task Render<TModel>(TModel model, Stream output, Cancel cancel = default) =>
        Render(model, output, null, cancel);

    /// <summary>
    /// Renders the template registered for <typeparamref name="TModel"/> and stamps the document's
    /// properties. Only the values set on <paramref name="properties"/> are written; each part is
    /// merged, so properties the template carries of its own survive.
    /// </summary>
    /// <remarks>
    /// The model type is the whole address: a template is registered by its model type, or found
    /// through the definition the source generator's module initializer stored for it, so there is
    /// nothing further for the caller to name.
    /// </remarks>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public Task Render<TModel>(TModel model, Stream output, WordDocumentProperties? properties, Cancel cancel = default)
    {
        // Before the template is resolved: for a generated template the first render is what
        // scans and validates it.
        cancel.ThrowIfCancellationRequested();

        var template = Resolve(typeof(TModel));

        if (model == null)
        {
            throw new ParchmentRenderException(
                typeof(TModel).Name,
                $"Model is null. Template is registered for {template.ModelType.Name}; pass a non-null instance.");
        }

        return template.Render(model, output, properties, cancel);
    }

    /// <inheritdoc cref="Render{TModel}(TModel, Stream, WordDocumentProperties, Cancel)"/>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public Task RenderToFile<TModel>(TModel model, string path, Cancel cancel = default) =>
        RenderToFile(model, path, null, cancel);

    /// <inheritdoc cref="Render{TModel}(TModel, Stream, WordDocumentProperties, Cancel)"/>
    /// <remarks>
    /// The file at <paramref name="path"/> is only replaced once the render has succeeded. A
    /// failed or cancelled render leaves an existing file as it was and creates nothing at a new
    /// path.
    /// </remarks>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public async Task RenderToFile<TModel>(TModel model, string path, WordDocumentProperties? properties, Cancel cancel = default)
    {
        // Before the file is created, so an already cancelled call leaves nothing at the path.
        cancel.ThrowIfCancellationRequested();

        // Rendered into a sibling temporary file and moved into place, so the destination is
        // never truncated ahead of a render that then throws, and never holds half a document
        // when the write itself fails. The same directory keeps the move on one volume, where it
        // is a rename rather than a copy.
        ProbeDestination(path);
        var directory = Path.GetDirectoryName(Path.GetFullPath(path))!;
        var temp = Path.Combine(directory, Path.GetRandomFileName());
        var file = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        try
        {
            await using (file.ConfigureAwait(false))
            {
                await Render(model, file, properties, cancel).ConfigureAwait(false);
            }

            File.Move(temp, path, true);
        }
        catch
        {
            DeleteQuietly(temp);
            throw;
        }
    }

    // Opens an existing destination the way File.Create would, without truncating it, so a missing
    // directory, a locked file or a read-only one fails before the render and with the exception
    // writing to the path directly gives. Replacing a locked file by a move reports
    // UnauthorizedAccessException on Windows where opening it reports IOException.
    static void ProbeDestination(string path)
    {
        try
        {
            using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
            }
        }
        catch (FileNotFoundException)
        {
        }
    }

    // Cleanup on a path that is already failing: the render's own exception is the one to surface.
    static void DeleteQuietly(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>
    /// The registered template for <paramref name="modelType"/>, materializing the source
    /// generator's definition on first use.
    /// </summary>
    /// <remarks>
    /// A definition holds raw content only — the generator's module initializer stores bytes and
    /// text, and the scan/validate work runs here, once per store, under this store's policies.
    /// </remarks>
    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    internal RegisteredTemplate Resolve(Type modelType)
    {
        // Walk the inheritance chain so a template registered against an abstract base renders a
        // concrete subclass instance — the documented polymorphic binding surface. Nearest type
        // wins.
        for (var type = modelType; type != null && type != typeof(object); type = type.BaseType)
        {
            if (templates.TryGetValue(type, out var existing))
            {
                return existing;
            }

            if (GeneratedTemplateDefinitions.TryGet(type, out var definition))
            {
                lock (materializeLock)
                {
                    if (!templates.TryGetValue(type, out existing))
                    {
                        Materialize(type, definition);
                        existing = templates[type];
                    }
                }

                return existing;
            }
        }

        throw new ParchmentRenderException(modelType.Name, $"No template is registered for {modelType.Name}");
    }

    [RequiresUnreferencedCode(Unsupported.Trimming)]
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    void Materialize(Type modelType, TemplateDefinition definition)
    {
        switch (definition)
        {
            case DocxTemplateDefinition docx:
                RegisterDocxTemplate(modelType, DocxCloner.ToWritableStream(docx.Template()), docx.Protection);
                break;
            case MarkdownTemplateDefinition markdown:
                RegisterMarkdownTemplate(modelType, markdown.Markdown, markdown.StyleSource?.Invoke());
                break;
            default:
                throw new InvalidOperationException($"Unknown template definition: {definition.GetType().Name}");
        }
    }

    public static void AddFilter(string name, FilterDelegate filter) =>
        SharedFluid.AddFilter(name, filter);
}
