/// <summary>
/// Walks a cached scope tree against a cloned docx, evaluating substitution tokens and block tags
/// to produce the final rendered document.
/// </summary>
class ScopeTreeRunner(
    string templateName,
    string partUri,
    Dictionary<string, Paragraph> anchorMap,
    TemplateContext context,
    MainDocumentPart mainPart,
    object rootModel,
    ExcelsiorTableMap excelsiorTables,
    FormatMap formats,
    StringListMap stringLists,
    EditableMap editables,
    EditableState editableState,
    WordNumberingState numberingState,
    Lazy<StyleSet> styles,
    ImagePolicies imagePolicies)
{
    List<StructuralReplacement>? structuralReplacements;

    public Task RunAsync(IReadOnlyList<RangeNode> nodes) =>
        ProcessAsync(nodes);

    public void ApplyStructural()
    {
        if (structuralReplacements == null)
        {
            return;
        }

        foreach (var replacement in structuralReplacements)
        {
            ApplyOne(replacement);
        }

        structuralReplacements.Clear();
    }

    static void ApplyOne(StructuralReplacement replacement)
    {
        var host = replacement.Host;
        var parent = host.Parent;
        if (parent == null)
        {
            return;
        }

        OpenXmlElement cursor = host;
        foreach (var produced in replacement.Produced)
        {
            cursor = parent.InsertAfter(produced, cursor);
        }

        host.Remove();
    }

    async Task ProcessAsync(IReadOnlyList<RangeNode> nodes)
    {
        foreach (var node in nodes)
        {
            await ProcessAsync(node);
        }
    }

    Task ProcessAsync(RangeNode node) =>
        node switch
        {
            SubstitutionNode substitution => ProcessSubstitutionAsync(substitution),
            LoopNode loop => ProcessLoopAsync(loop),
            IfNode ifNode => ProcessIfAsync(ifNode),
            _ => Task.CompletedTask
        };

    async Task ProcessSubstitutionAsync(SubstitutionNode node)
    {
        if (!anchorMap.TryGetValue(node.AnchorName, out var host))
        {
            return;
        }

        // Cache the ParagraphText across token replacements. Tokens are applied in reverse-offset
        // order, so a higher-offset Replace doesn't disturb the spans/offsets used for the next
        // lower-offset token. Structural ops (splice/split) and MutateToken arbitrary mutation do
        // invalidate the cache — clear it after those.
        ParagraphText? cachedText = null;
        ParagraphText Text() => cachedText ??= ParagraphText.Build(host);

        // Snapshot the host paragraph's original character length once.
        var originalLength = Text().InnerText.Length;

        // node.Tokens come from TokenScan in ascending-offset order. Iterate in reverse without
        // allocating an OrderByDescending().ToList() — applying high-offset replacements first
        // keeps the lower-offset spans valid for the next iteration.
        List<(DocxTokenSite site, object value)>? soloStructuralTokens = null;
        var splitQueued = false;
        var tokenCount = node.Tokens.Count;
        var hasEditableSibling = HasEditableToken(node);

        for (var i = tokenCount - 1; i >= 0; i--)
        {
            var token = node.Tokens[i];
            var evaluated = await EvaluateTokenAsync(token, host, tokenCount);
            if (evaluated is EditableToken editable)
            {
                if (editable.Entry.Kind == EditableFieldKind.Html)
                {
                    QueueEditableHtml(host, token, editable, tokenCount, originalLength);
                    continue;
                }

                EditableSplicer.Insert(
                    host,
                    token.Offset,
                    token.Length,
                    sitePr => EditableFieldBuilder.Build(editable.Entry, editable.Value, sitePr, editableState, context.CultureInfo));
                cachedText = null;
                continue;
            }

            if (evaluated is MarkdownToken or HtmlToken or OpenXmlToken)
            {
                if (hasEditableSibling)
                {
                    // Structural splice/split clones paragraph halves, which would corrupt the
                    // editable field's perm-range markers. Statically-known cases are rejected
                    // at registration; this guards TokenValue-typed properties that only turn
                    // out structural at render time.
                    throw new ParchmentRenderException(
                        templateName,
                        $"Token '{token.Source}' produced structural content in a paragraph that also contains an editable field. Move one of them to its own paragraph.",
                        partUri,
                        Snippet(host, token),
                        token.Source);
                }

                if (tokenCount == 1 && token.Offset == 0 && token.Length == originalLength)
                {
                    // Whole host paragraph is the token — queue for replacement after every
                    // other in-paragraph substitution has run.
                    (soloStructuralTokens ??= []).Add((token, evaluated));
                    continue;
                }

                ApplyNonSoloStructural(token, host, evaluated, ref splitQueued);
                cachedText = null;
                continue;
            }

            if (evaluated is MutateToken mutate)
            {
                // Clear the token text, then hand the paragraph to the caller for in-place mutation.
                Text().Replace(token.Offset, token.Length, string.Empty);
                var ctx = new OpenXmlContextImpl(mainPart, numberingState, styles.Value, host);
                mutate.Apply(host, ctx);
                cachedText = null;
                continue;
            }

            var replacement = ToDisplayString(evaluated);
            Text().Replace(token.Offset, token.Length, replacement);
        }

        if (soloStructuralTokens is { Count: > 0 })
        {
            (structuralReplacements ??= []).Add(new(host, BuildStructuralReplacements(host, soloStructuralTokens)));
        }
    }

    /// <summary>
    /// Queues an editable-HTML field as a block-level structural replacement. The member renders a
    /// rich-content control the user can edit within an editable range; it must occupy its own
    /// paragraph (like a read-only [Html] token) so the block <c>w:sdt</c> + perm range replace the
    /// host cleanly. Extraction serializes the block content back to HTML.
    /// </summary>
    void QueueEditableHtml(Paragraph host, DocxTokenSite token, EditableToken editable, int tokenCount, int originalLength)
    {
        if (tokenCount != 1 ||
            token.Offset != 0 ||
            token.Length != originalLength)
        {
            throw new ParchmentRenderException(
                templateName,
                $"Editable HTML token '{token.Source}' must sit alone in its paragraph — it renders a block-level editable region. Move it to its own paragraph.",
                partUri,
                Snippet(host, token),
                token.Source);
        }

        var html = editable.Value as string;
        IReadOnlyList<OpenXmlElement> content = string.IsNullOrWhiteSpace(html)
            ? []
            : WordHtmlConverter.ToElements(
                html,
                mainPart,
                imagePolicies.BuildSettings(numberingSession: numberingState.GetHtmlSession()));

        (structuralReplacements ??= []).Add(new(host, EditableHtmlBuilder.Build(editable.Entry, content, editableState, host)));
    }

    /// <summary>
    /// Splice or split the host paragraph for a non-solo structural token. Inline-equivalent
    /// output (single produced paragraph) is unwrapped and spliced in place; anything else
    /// (multiple blocks, a table) splits the host paragraph at the token offset and inserts
    /// the produced elements between the two halves.
    /// </summary>
    void ApplyNonSoloStructural(DocxTokenSite token, Paragraph host, object value, ref bool splitQueued)
    {
        var produced = RenderTokenValue(value, host);
        if (produced.Count == 0)
        {
            // Nothing to render — strip the token text from the host paragraph.
            ParagraphText.Build(host).Replace(token.Offset, token.Length, string.Empty);
            return;
        }

        if (ParagraphSplicer.IsInlineEquivalent(produced))
        {
            ParagraphSplicer.SpliceInline(host, token.Offset, token.Length, (Paragraph)produced[0]);
            return;
        }

        if (splitQueued)
        {
            // A second block-shaped substitution on the same paragraph would create overlapping
            // structural replacements; we don't try to compose them. The author needs to give
            // the second token its own paragraph.
            throw new ParchmentRenderException(
                templateName,
                $"Token '{token.Source}' produced block-level content but another structural substitution on the same paragraph already required a paragraph split. Move one of the tokens to its own paragraph.",
                partUri,
                Snippet(host, token),
                token.Source);
        }

        // Block-shaped output in a non-solo context: split the host paragraph and insert the
        // produced block elements between the resulting before/after halves. We queue this as a
        // structural replacement so any other in-paragraph substitutions on the same host have
        // already applied to the host's text by the time we replace it.
        var split = ParagraphSplicer.Split(host, token.Offset, token.Length, produced);
        (structuralReplacements ??= []).Add(new(host, split));
        splitQueued = true;
    }

    List<OpenXmlElement> BuildStructuralReplacements(Paragraph host, IReadOnlyList<(DocxTokenSite site, object value)> tokens)
    {
        var result = new List<OpenXmlElement>();
        foreach (var (_, value) in tokens)
        {
            result.AddRange(RenderTokenValue(value, host));
        }

        return result;
    }

    IReadOnlyList<OpenXmlElement> RenderTokenValue(object value, Paragraph host) =>
        value switch
        {
            MarkdownToken md => MarkdownRendering.Render(md.Source, mainPart, numberingState, imagePolicies, headingOffset: 0),
            HtmlToken html => WordHtmlConverter.ToElements(
                html.Source,
                mainPart,
                imagePolicies.BuildSettings(numberingSession: numberingState.GetHtmlSession())),
            OpenXmlToken raw when ReferenceEquals(raw, OpenXmlToken.Empty) => [],
            OpenXmlToken raw => raw
                .Render(new OpenXmlContextImpl(mainPart, numberingState, styles.Value, host))
                .ToList(),
            _ => []
        };

    /// <summary>
    /// Splits the synchronous fast path from async machinery: the dispatch checks (Excelsior /
    /// Format / StringList) are pure-sync, and the dominant Fluid path (MemberExpression member
    /// access on a POCO) completes synchronously. Stays out of an async state machine entirely
    /// when EvaluateAsync returns a synchronously-completed ValueTask. Saves the per-token Task
    /// allocation on the hot loop-body path.
    /// </summary>
    ValueTask<object> EvaluateTokenAsync(DocxTokenSite site, Paragraph host, int siblingCount)
    {
        try
        {
            if (TryResolveExcelsiorTable(site) is { } excelsiorToken)
            {
                return new(excelsiorToken);
            }

            if (TryResolveFormatted(site) is { } formatted)
            {
                return new(formatted);
            }

            if (TryResolveEditable(site) is { } editable)
            {
                return new(editable);
            }

            if (TryResolveStringList(site, host, siblingCount) is { } stringList)
            {
                return new(stringList);
            }

            // Walk the parsed FluidTemplate to its OutputStatement and evaluate the underlying
            // Expression directly (filter chain included). This lets us see whether the value is a
            // TokenValue (markdown / openxml hatch) before falling back to string rendering, without
            // round-tripping through the Render() pipeline twice.
            //
            // DocxTokenSite.Template is only ever constructed from a `{{ ... }}` substitution body
            // by TokenScanner.ParseSubstitution — FluidParser emits that as exactly one
            // OutputStatement. Block tags go down a separate path and never reach here.
            var statements = ((Fluid.Parser.FluidTemplate)site.Template).Statements;
            var output = (OutputStatement)statements[0];
            var pending = output.Expression.EvaluateAsync(context);
            if (pending.IsCompletedSuccessfully)
            {
                return new(InterpretFluidValue(pending.Result));
            }

            return AwaitFluidValue(pending, host, site);
        }
        catch (ParchmentException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw WrapException(exception, host, site);
        }
    }

    static object InterpretFluidValue(FluidValue fluidValue)
    {
        // TokenValue is always wrapped in ObjectValue (see Filters.Markdown / BulletList /
        // NumberedList and TokenValueHelpers). Skip ToObjectValue for primitive FluidValues
        // (StringValue, NumberValue, BooleanValue, ArrayValue, DateTimeValue, ...) — that
        // path boxes numerics to object before the type test, which the common-case
        // text-token path doesn't need.
        if (fluidValue is ObjectValue &&
            fluidValue.ToObjectValue() is TokenValue tokenValue)
        {
            return tokenValue;
        }

        return fluidValue.ToStringValue();
    }

    async ValueTask<object> AwaitFluidValue(ValueTask<FluidValue> pending, Paragraph host, DocxTokenSite site)
    {
        try
        {
            return InterpretFluidValue(await pending);
        }
        catch (ParchmentException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw WrapException(exception, host, site);
        }
    }

    ParchmentRenderException WrapException(Exception exception, Paragraph host, DocxTokenSite site) =>
        new(
            templateName,
            exception.Message,
            partUri,
            Snippet(host, site),
            site.Source,
            inner: exception);

    OpenXmlToken? TryResolveExcelsiorTable(DocxTokenSite site)
    {
        if (excelsiorTables.IsEmpty ||
            site.References.Count == 0)
        {
            return null;
        }

        // Match the token's full dotted reference (e.g. `Customer.Lines`) against the registered
        // map. Single-segment, multi-segment, and arbitrarily-nested paths from the root model
        // all flow through the same lookup. Loop-scope variables (e.g. `{{ line.SubItems }}`
        // inside `{% for line in Lines %}`) won't match because the map is keyed on paths from
        // the root model only — they fall through to normal Fluid evaluation.
        var reference = site.References[0];
        if (!excelsiorTables.TryGet(reference.Dotted, out var entry))
        {
            return null;
        }

        var data = entry.Getter(rootModel);
        if (data == null)
        {
            return OpenXmlToken.Empty;
        }

        return new(_ => [ExcelsiorTableBridge.BuildTable(entry.ElementType, data, mainPart, entry.HeadingParagraphStyle, entry.BodyParagraphStyle)]);
    }

    TokenValue? TryResolveStringList(DocxTokenSite site, Paragraph host, int siblingCount)
    {
        if (stringLists.IsEmpty ||
            site.References.Count == 0)
        {
            return null;
        }

        var reference = site.References[0];
        if (!stringLists.TryGet(reference.Dotted, out var getter))
        {
            return null;
        }

        // Auto-bullet rendering swaps the entire host paragraph. Skip silently if the token
        // doesn't sit alone — the user gets Fluid stringification in that case (consistent with
        // pre-feature behavior) instead of a surprising paragraph swap that drops surrounding text.
        if (siblingCount != 1)
        {
            return null;
        }

        var paragraphText = ParagraphText.Build(host).InnerText;
        if (site.Offset != 0 ||
            site.Length != paragraphText.Length)
        {
            return null;
        }

        // If the user attached a filter chain (e.g. `{{ Tags | numbered_list }}`), they're
        // explicitly opting into Fluid-driven rendering — let that path handle it.
        var statements = ((Fluid.Parser.FluidTemplate)site.Template).Statements;
        if (statements.Count == 0 ||
            statements[0] is not OutputStatement { Expression: MemberExpression })
        {
            return null;
        }

        var data = getter(rootModel);
        if (data is not IEnumerable<string> items)
        {
            return OpenXmlToken.Empty;
        }

        // Materialize so the deferred render delegate doesn't re-enumerate a fresh sequence
        // (and so a null model walk can't surface here).
        return TokenValueHelpers.BulletList(items.ToList());
    }

    EditableToken? TryResolveEditable(DocxTokenSite site)
    {
        if (editables.IsEmpty ||
            site.References.Count == 0)
        {
            return null;
        }

        if (!editables.TryGet(site.References[0].Dotted, out var entry))
        {
            return null;
        }

        return new(entry, entry.Getter(rootModel));
    }

    bool HasEditableToken(SubstitutionNode node)
    {
        if (editables.IsEmpty ||
            node.Tokens.Count < 2)
        {
            // A solo structural token never conflicts with an editable sibling.
            return false;
        }

        foreach (var token in node.Tokens)
        {
            if (token.References.Count > 0 &&
                editables.TryGet(token.References[0].Dotted, out _))
            {
                return true;
            }
        }

        return false;
    }

    TokenValue? TryResolveFormatted(DocxTokenSite site)
    {
        if (formats.IsEmpty || site.References.Count == 0)
        {
            return null;
        }

        var reference = site.References[0];
        if (!formats.TryGet(reference.Dotted, out var entry))
        {
            return null;
        }

        var value = entry.Getter(rootModel);
        var text = value as string ?? string.Empty;
        return entry.Kind switch
        {
            FormatMapKind.Html => new HtmlToken(text),
            FormatMapKind.Markdown => new MarkdownToken(text),
            _ => null
        };
    }

    async Task ProcessLoopAsync(LoopNode loop)
    {
        if (!anchorMap.TryGetValue(loop.OpenAnchorName, out var open) ||
            !anchorMap.TryGetValue(loop.CloseAnchorName, out var close))
        {
            return;
        }

        var parent = open.Parent;
        if (parent == null)
        {
            return;
        }

        var items = await ResolveIterableAsync(loop);
        var bodyElements = CaptureBetween(open, close);

        // A loop over an [EditableField] collection renders as an editable Word repeating section
        // rather than the read-only inline expansion. Body-only: header/footer parts get
        // EditableMap.Empty, so TryGetCollection is false there and the loop stays read-only.
        var sourceRefs = IdentifierVisitor.Collect(loop.LoopSource);
        if (sourceRefs.Count > 0 &&
            editables.TryGetCollection(sourceRefs[0].Dotted, out var collectionEntry))
        {
            await RenderEditableCollection(loop, collectionEntry, open, parent, bodyElements, items.ToList());
            foreach (var element in bodyElements)
            {
                element.Remove();
            }

            open.Remove();
            close.Remove();
            return;
        }

        OpenXmlElement insertAnchor = open;

        var cloneAnchors = new Dictionary<string, Paragraph>(StringComparer.Ordinal);
        var pendingMoves = new List<OpenXmlElement>();
        // Reuse a single scratch parent across iterations. After each iteration's child-move
        // loop the scratch is empty again, so plain reuse is safe.
        var scratch = new Body();

        // Construct the inner runner once and reuse across iterations. cloneAnchors is passed
        // by reference and re-populated per iteration, so the runner sees the fresh per-iteration
        // anchor map without needing to be re-allocated. ApplyStructural clears the runner's
        // structuralReplacements list at the end of each iteration.
        // Loop bodies deliberately get an empty editable map: a root-pathed editable token
        // inside {% for %} would stamp one content control per iteration, producing duplicate
        // tags that break extraction. Loop-body editable tokens render as plain read-only text.
        var clonedRunner = new ScopeTreeRunner(
            templateName,
            partUri,
            cloneAnchors,
            context,
            mainPart,
            rootModel,
            excelsiorTables,
            formats,
            stringLists,
            EditableMap.Empty,
            editableState,
            numberingState,
            styles,
            imagePolicies);

        foreach (var item in items)
        {
            // Each iteration gets its own non-bubbling child scope so the loop variable binding
            // is confined to the iteration. Without this, context.SetValue lands in the root
            // scope and the binding leaks past `{% endfor %}` — shadowing model properties of
            // the same name and (in nested-same-name loops) overwriting the outer's binding.
            // Parchment block tags are limited to for/if/elsif/else/endif (no `{% assign %}`),
            // so the non-bubbling EnterChildScope is preferred over EnterForLoopScope which
            // exists to let assigns inside a loop persist after it.
            context.EnterChildScope();
            try
            {
                context.SetValue(loop.LoopVariable, item);
                cloneAnchors.Clear();

                // Clone each body element directly into the scratch parent. Original elements stay
                // attached to the live document (removed below after the loop completes) and we
                // never mutate them, so cloning from them is equivalent to cloning from a detached
                // template. Attaching to scratch is required because nested ProcessLoopAsync /
                // ProcessIfAsync / ApplyStructural rely on Parent and sibling traversal — on
                // detached clones those return null and inner scope tags silently no-op.
                foreach (var element in bodyElements)
                {
                    var clone = element.CloneNode(true);
                    scratch.AppendChild(clone);
                    CollectAnchors(clone, cloneAnchors);
                }

                // Reuse the original scope tree directly. cloneAnchors is keyed on the same anchor
                // names the registration-time tree references, so no per-iteration tree rebuild
                // (Remap) is needed. Multiple iterations leave duplicate-named bookmarks in the
                // live docx — that's fine because StripAll runs before Save and removes them all.
                await clonedRunner.RunAsync(loop.Body);
                clonedRunner.ApplyStructural();

                // Snapshot scratch's children before mutating, so remove+insert doesn't break
                // the live ChildElements enumeration.
                pendingMoves.Clear();
                foreach (var element in scratch.ChildElements)
                {
                    pendingMoves.Add(element);
                }

                foreach (var produced in pendingMoves)
                {
                    produced.Remove();
                    insertAnchor = parent.InsertAfter(produced, insertAnchor);
                }
            }
            finally
            {
                context.ReleaseScope();
            }
        }

        foreach (var element in bodyElements)
        {
            element.Remove();
        }

        open.Remove();
        close.Remove();
    }

    /// <summary>
    /// Renders an editable-collection loop as a Word repeating section: one
    /// <c>w15:repeatingSectionItem</c> per model item (each rendered with an item-scoped editable map
    /// so <c>{{ item.Field }}</c> tokens become item-relative controls), wrapped in the repeating-section
    /// container + editable perm range. An empty collection still renders one blank item so Word has a
    /// clone template.
    /// </summary>
    async Task RenderEditableCollection(
        LoopNode loop,
        CollectionEntry entry,
        OpenXmlElement open,
        OpenXmlElement parent,
        List<OpenXmlElement> bodyElements,
        List<FluidValue> items)
    {
        var effectiveItems = items.Count > 0
            ? items
            : [FluidValue.Create(entry.ElementFactory(), context.Options)];

        var scratch = new Body();
        var cloneAnchors = new Dictionary<string, Paragraph>(StringComparer.Ordinal);
        var repeatedItems = new List<SdtBlock>();

        foreach (var fluidItem in effectiveItems)
        {
            var rawItem = fluidItem.ToObjectValue();
            context.EnterChildScope();
            try
            {
                context.SetValue(loop.LoopVariable, fluidItem);
                cloneAnchors.Clear();
                foreach (var element in bodyElements)
                {
                    var clone = element.CloneNode(true);
                    scratch.AppendChild(clone);
                    CollectAnchors(clone, cloneAnchors);
                }

                var itemRunner = new ScopeTreeRunner(
                    templateName,
                    partUri,
                    cloneAnchors,
                    context,
                    mainPart,
                    rootModel,
                    excelsiorTables,
                    formats,
                    stringLists,
                    entry.ElementMap.ScopedToItem(loop.LoopVariable, rawItem),
                    editableState,
                    numberingState,
                    styles,
                    imagePolicies);
                await itemRunner.RunAsync(loop.Body);
                itemRunner.ApplyStructural();

                var itemBody = new List<OpenXmlElement>();
                foreach (var element in scratch.ChildElements.ToList())
                {
                    element.Remove();
                    itemBody.Add(element);
                }

                repeatedItems.Add(EditableRepeatingBuilder.BuildItem(itemBody, editableState));
            }
            finally
            {
                context.ReleaseScope();
            }
        }

        var anchor = open;
        foreach (var produced in EditableRepeatingBuilder.BuildContainer(entry, repeatedItems, editableState))
        {
            anchor = parent.InsertAfter(produced, anchor);
        }
    }

    static List<OpenXmlElement> CaptureBetween(OpenXmlElement start, OpenXmlElement end)
    {
        var result = new List<OpenXmlElement>();
        var cursor = start.NextSibling();
        while (cursor != null &&
               cursor != end)
        {
            result.Add(cursor);
            cursor = cursor.NextSibling();
        }

        return result;
    }

    /// <summary>
    /// Index every Parchment-prefixed BookmarkStart in <paramref name="clone"/> by its existing
    /// (registration-time) name → host paragraph. The unchanged scope tree references those
    /// same names, so the runner can look up the cloned host directly without a name-rewrite
    /// or remap pass. Duplicate names across iterations are harmless because each iteration
    /// uses its own fresh anchor dictionary, and StripAll prefix-matches at save time.
    /// </summary>
    static void CollectAnchors(OpenXmlElement clone, Dictionary<string, Paragraph> anchorMap)
    {
        // Parchment-prefixed bookmarks are always direct children of a Paragraph (see
        // Anchors.InsertAfterProperties). Walk Paragraphs and scan their direct children
        // instead of Descendants<BookmarkStart>() over the full subtree — the descendant
        // enumerator allocates and walks every non-bookmark element, while direct-child
        // iteration touches only ParagraphProperties + bookmark + run-level descendants of
        // the Paragraph element directly.
        if (clone is Paragraph paragraph)
        {
            CollectFromParagraph(paragraph, anchorMap);
            return;
        }

        foreach (var p in clone.Descendants<Paragraph>())
        {
            CollectFromParagraph(p, anchorMap);
        }
    }

    static void CollectFromParagraph(Paragraph paragraph, Dictionary<string, Paragraph> anchorMap)
    {
        foreach (var child in paragraph.ChildElements)
        {
            if (child is BookmarkStart {Name.Value: { } name} &&
                name.StartsWith(Anchors.Prefix, StringComparison.Ordinal))
            {
                anchorMap[name] = paragraph;
            }
        }
    }

    async Task<IEnumerable<FluidValue>> ResolveIterableAsync(LoopNode loop)
    {
        // Hand the loop source straight to Fluid: ForStatement.Source is an Expression that, when
        // evaluated, yields a FluidValue we can enumerate. This honors filters, complex paths,
        // arithmetic, and any value converters Fluid is configured with — none of which the previous
        // reflection-walk supported.
        var sourceValue = await loop.LoopSource.EvaluateAsync(context);
        return sourceValue.Enumerate(context);
    }

    async Task ProcessIfAsync(IfNode ifNode)
    {
        if (!anchorMap.TryGetValue(ifNode.OpenAnchorName, out var open) ||
            !anchorMap.TryGetValue(ifNode.CloseAnchorName, out var close))
        {
            return;
        }

        if (open.Parent == null)
        {
            return;
        }

        IReadOnlyList<RangeNode>? chosen = null;
        var chosenIndex = -1;
        for (var i = 0; i < ifNode.Branches.Count; i++)
        {
            if (!await EvaluateConditionAsync(ifNode.Branches[i].Condition))
            {
                continue;
            }

            chosen = ifNode.Branches[i].Body;
            chosenIndex = i;
            break;
        }

        // Fall back to the else branch when it exists. Signalled by the tag anchor, not
        // ElseBody.Count — a static-only else branch has an empty body (static paragraphs carry
        // no tree nodes) but still owns a physical range that must render.
        if (chosen == null &&
            ifNode.ElseAnchorName != null)
        {
            chosen = ifNode.ElseBody;
        }

        // Everything physically between {% if %} and {% endif %}: branch tag paragraphs, token
        // paragraphs, static paragraphs, tables.
        var allBranchElements = CaptureBetween(open, close);

        if (chosen == null)
        {
            foreach (var element in allBranchElements)
            {
                element.Remove();
            }

            open.Remove();
            close.Remove();
            return;
        }

        // The chosen branch's content is kept positionally — everything strictly between the
        // branch's tag paragraph and the next boundary tag ({% elsif %} / {% else %} /
        // {% endif %}). An anchor-derived keep-set would drop static paragraphs and tables:
        // neither carries an anchor or a scope-tree node.
        var (start, end) = ChosenBranchBoundaries(ifNode, chosenIndex, open, close);
        var keepElements = CaptureBetween(start, end);

        // Process the chosen branch in place (no cloning — branch elements are used once).
        // Anchors are collected across nested content too, so tokens inside tables resolve.
        var branchAnchors = new Dictionary<string, Paragraph>(StringComparer.Ordinal);
        foreach (var element in keepElements)
        {
            CollectAnchors(element, branchAnchors);
        }

        var innerRunner = new ScopeTreeRunner(templateName, partUri, branchAnchors, context, mainPart, rootModel, excelsiorTables, formats, stringLists, editables, editableState, numberingState, styles, imagePolicies);
        await innerRunner.RunAsync(chosen);
        innerRunner.ApplyStructural();

        var keep = new HashSet<OpenXmlElement>(keepElements);
        foreach (var element in allBranchElements)
        {
            if (!keep.Contains(element))
            {
                element.Remove();
            }
        }

        open.Remove();
        close.Remove();
    }

    /// <summary>
    /// The paragraphs bounding the chosen branch's physical content: its own tag paragraph and
    /// the next branch boundary (the following <c>{% elsif %}</c>, the <c>{% else %}</c>, or
    /// <c>{% endif %}</c>). <paramref name="chosenIndex"/> of -1 means the else branch.
    /// </summary>
    (Paragraph Start, Paragraph End) ChosenBranchBoundaries(IfNode ifNode, int chosenIndex, Paragraph open, Paragraph close)
    {
        if (chosenIndex < 0)
        {
            var elseStart = anchorMap.GetValueOrDefault(ifNode.ElseAnchorName!, open);
            return (elseStart, close);
        }

        var start = anchorMap.GetValueOrDefault(ifNode.Branches[chosenIndex].TagAnchorName, open);

        if (chosenIndex + 1 < ifNode.Branches.Count &&
            anchorMap.TryGetValue(ifNode.Branches[chosenIndex + 1].TagAnchorName, out var nextBranchTag))
        {
            return (start, nextBranchTag);
        }

        if (ifNode.ElseAnchorName != null &&
            anchorMap.TryGetValue(ifNode.ElseAnchorName, out var elseTag))
        {
            return (start, elseTag);
        }

        return (start, close);
    }

    async Task<bool> EvaluateConditionAsync(Expression condition) =>
        (await condition.EvaluateAsync(context)).ToBooleanValue();

    static string ToDisplayString(object value) =>
        value switch
        {
            null => string.Empty,
            string s => s,
            _ => value.ToString() ?? string.Empty
        };

    static string Snippet(Paragraph host, DocxTokenSite site)
    {
        var text = ParagraphText.Build(host).InnerText;
        var start = Math.Max(0, site.Offset - 40);
        var end = Math.Min(text.Length, site.Offset + site.Length + 40);
        return text[start..end];
    }

    sealed record StructuralReplacement(Paragraph Host, IReadOnlyList<OpenXmlElement> Produced);
}
