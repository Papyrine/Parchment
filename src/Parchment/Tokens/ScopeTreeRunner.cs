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

        for (var i = tokenCount - 1; i >= 0; i--)
        {
            var token = node.Tokens[i];
            var evaluated = await EvaluateTokenAsync(token, host, tokenCount);
            if (evaluated is MarkdownToken or HtmlToken or OpenXmlToken)
            {
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

            if (TryResolveStringList(site, host, siblingCount) is { } stringList)
            {
                return new(stringList);
            }

            // Walk the parsed FluidTemplate to its OutputStatement and evaluate the underlying
            // Expression directly (filter chain included). This lets us see whether the value is a
            // TokenValue (markdown / openxml hatch) before falling back to string rendering, without
            // round-tripping through the Render() pipeline twice.
            var statements = ((Fluid.Parser.FluidTemplate)site.Template).Statements;
            if (statements.Count > 0 &&
                statements[0] is OutputStatement output)
            {
                var pending = output.Expression.EvaluateAsync(context);
                if (pending.IsCompletedSuccessfully)
                {
                    return new(InterpretFluidValue(pending.Result));
                }

                return AwaitFluidValue(pending, host, site);
            }

            return RenderFullTemplateAsync(site, host);
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

    async ValueTask<object> RenderFullTemplateAsync(DocxTokenSite site, Paragraph host)
    {
        try
        {
            await using var writer = new StringWriter();
            await site.Template.RenderAsync(writer, System.Text.Encodings.Web.HtmlEncoder.Default, context);
            return writer.ToString();
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

        return new(_ => [ExcelsiorTableBridge.BuildTable(entry.ElementType, data, mainPart)]);
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

        var parent = open.Parent;
        if (parent == null)
        {
            return;
        }

        IReadOnlyList<RangeNode>? chosen = null;
        foreach (var branch in ifNode.Branches)
        {
            if (!await EvaluateConditionAsync(branch.Condition))
            {
                continue;
            }

            chosen = branch.Body;
            break;
        }

        if (chosen == null &&
            ifNode.ElseBody.Count > 0)
        {
            chosen = ifNode.ElseBody;
        }

        // Collect all branch paragraphs between open and close — everything that should be removed
        var allBranchParagraphs = CaptureBetween(open, close);

        if (chosen == null)
        {
            foreach (var element in allBranchParagraphs)
            {
                element.Remove();
            }
        }
        else
        {
            // Process chosen branch in place (no cloning — branch paragraphs are used once)
            var branchAnchors = new Dictionary<string, Paragraph>(StringComparer.Ordinal);
            foreach (var p in allBranchParagraphs.OfType<Paragraph>())
            {
                var start = p.Elements<BookmarkStart>()
                    .FirstOrDefault(_ => _.Name?.Value?.StartsWith(Anchors.Prefix, StringComparison.Ordinal) == true);
                if (start?.Name?.Value != null)
                {
                    branchAnchors[start.Name.Value] = p;
                }
            }

            var innerRunner = new ScopeTreeRunner(templateName, partUri, branchAnchors, context, mainPart, rootModel, excelsiorTables, formats, stringLists, numberingState, styles, imagePolicies);
            await innerRunner.RunAsync(chosen);
            innerRunner.ApplyStructural();

            // Only keep paragraphs that belong to the chosen branch; remove others
            var keep = new HashSet<OpenXmlElement>(CollectBranchParagraphs(chosen, branchAnchors));
            foreach (var element in allBranchParagraphs)
            {
                if (!keep.Contains(element))
                {
                    element.Remove();
                }
            }
        }

        open.Remove();
        close.Remove();
    }

    static IEnumerable<OpenXmlElement> CollectBranchParagraphs(
        IReadOnlyList<RangeNode> nodes,
        Dictionary<string, Paragraph> anchors)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case SubstitutionNode s when anchors.TryGetValue(s.AnchorName, out var p):
                    yield return p;
                    break;
                case StaticNode s when anchors.TryGetValue(s.AnchorName, out var p):
                    yield return p;
                    break;
                case LoopNode l:
                    if (anchors.TryGetValue(l.OpenAnchorName, out var lo))
                    {
                        yield return lo;
                    }

                    foreach (var child in CollectBranchParagraphs(l.Body, anchors))
                    {
                        yield return child;
                    }

                    if (anchors.TryGetValue(l.CloseAnchorName, out var lc))
                    {
                        yield return lc;
                    }

                    break;
                case IfNode ifn:
                    if (anchors.TryGetValue(ifn.OpenAnchorName, out var io))
                    {
                        yield return io;
                    }

                    foreach (var branch in ifn.Branches)
                    {
                        foreach (var child in CollectBranchParagraphs(branch.Body, anchors))
                        {
                            yield return child;
                        }
                    }

                    foreach (var child in CollectBranchParagraphs(ifn.ElseBody, anchors))
                    {
                        yield return child;
                    }

                    if (anchors.TryGetValue(ifn.CloseAnchorName, out var ic))
                    {
                        yield return ic;
                    }

                    break;
            }
        }
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
