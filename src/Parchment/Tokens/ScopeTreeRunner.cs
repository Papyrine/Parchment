/// <summary>
/// Walks a cached scope tree against a cloned docx, evaluating substitution tokens and block tags
/// to produce the final rendered document.
/// </summary>
[RequiresUnreferencedCode(Unsupported.Trimming)]
[RequiresDynamicCode(Unsupported.DynamicCode)]
class ScopeTreeRunner(
    string templateName,
    string partUri,
    Dictionary<string, Paragraph> anchorMap,
    TemplateContext context,
    MainDocumentPart mainPart,
    PartRelationships relationships,
    object rootModel,
    ExcelsiorTableMap excelsiorTables,
    FormatMap formats,
    StringListMap stringLists,
    EditableMap editables,
    EditableState editableState,
    BookmarkIds bookmarks,
    CloneIds? cloneIds,
    WordNumberingState numberingState,
    Lazy<StyleSet> styles,
    ImagePolicies imagePolicies,
    Cancel cancel,
    List<Table>? tables = null,
    List<Paragraph>? sections = null,
    List<Paragraph>? carriedSections = null)
{
    List<StructuralReplacement>? structuralReplacements;

    // Every table that got a different block before it: one a structural replacement put at block
    // level, one that followed a block the render took out, one a loop iteration opens with.
    // Shared with the runners created for loop iterations and if branches. Whether one ends up
    // directly after another table is only known once the tag paragraphs around it are gone, so
    // the root runner checks at the end.
    readonly List<Table> tablesToCheck = tables ?? [];

    // Every [SECTION] marker a markdown value wrote, shared the same way. A marker holds only what
    // it declared, and the section it is derived from is whichever one it ends up inside - which a
    // loop iteration still sitting in its scratch container cannot say - so the root runner
    // finishes them at the end as well.
    readonly List<Paragraph> sectionMarkers = sections ?? [];

    // The paragraphs left standing in for a removed tag paragraph or token host that carried a
    // section break, shared the same way. Which paragraph ends the section instead is only known
    // once everything around it has been removed, repeated and replaced — see SettleSectionBreaks.
    readonly List<Paragraph> sectionBreaks = carriedSections ?? [];

    // The loop variables open around the nodes this runner walks. A token that starts with one
    // reads the loop item, so the root-keyed maps do not apply to it — see LoopScope.
    IReadOnlyList<string> loopVariables = [];

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

    /// <summary>
    /// Word merges two <c>w:tbl</c> siblings into one table. Run on the root runner once the whole
    /// tree has been processed: any table the render left directly after another table — one a
    /// token produced, the copy a loop iteration opens with, or a table of the template that lost
    /// the paragraphs keeping it apart — gets an empty paragraph before it. Two tables with
    /// nothing between them in the template are never recorded, and stay as authored.
    /// </summary>
    public void SeparateTables()
    {
        if (tablesToCheck.Count == 0)
        {
            return;
        }

        var recorded = new HashSet<OpenXmlElement>(tablesToCheck);
        var containers = new HashSet<OpenXmlElement>();
        foreach (var table in tablesToCheck)
        {
            // A table recorded and then removed with its branch or loop body has no parent.
            if (table.Parent is { } container &&
                containers.Add(container))
            {
                SeparateTablesIn(container, recorded);
            }
        }

        tablesToCheck.Clear();
    }

    // One pass over the container rather than asking each table for its previous sibling, which
    // the SDK answers by walking from the first child.
    static void SeparateTablesIn(OpenXmlElement container, HashSet<OpenXmlElement> recorded)
    {
        OpenXmlElement? previous = null;
        var afterTable = false;
        for (var child = container.FirstChild; child != null; child = child.NextSibling())
        {
            if (!IsRangeMarker(child))
            {
                if (afterTable &&
                    child is Table &&
                    recorded.Contains(child))
                {
                    container.InsertAfter(new Paragraph(), previous);
                }

                afterTable = child is Table;
            }

            previous = child;
        }
    }

    /// <summary>
    /// Turns each <c>[SECTION]</c> marker a markdown value wrote into a whole section, derived from
    /// the section of <paramref name="root"/> it landed in — see <see cref="SectionBreaks"/>. Run on
    /// the root runner once the whole tree has been processed and its replacements applied. Only a
    /// document body can hold a section break, so a marker anywhere else is rejected here.
    /// </summary>
    public void ApplySectionBreaks(OpenXmlCompositeElement root)
    {
        if (sectionMarkers.Count == 0)
        {
            return;
        }

        SectionBreaks.ApplyInPart(root, sectionMarkers, templateName, partUri);
        sectionMarkers.Clear();
    }

    // Range markers can sit between two blocks without keeping their tables apart.
    static bool IsRangeMarker(OpenXmlElement element) =>
        element is BookmarkStart or BookmarkEnd or PermStart or PermEnd or ProofError;

    static OpenXmlElement? PreviousBlock(OpenXmlElement element)
    {
        var cursor = element.PreviousSibling();
        while (cursor != null &&
               IsRangeMarker(cursor))
        {
            cursor = cursor.PreviousSibling();
        }

        return cursor;
    }

    /// <summary>
    /// Records the table that follows <paramref name="element"/>, when the render is about to
    /// change what comes before that table: by removing the element, or by having just put it
    /// there.
    /// </summary>
    void RecordTableAfter(OpenXmlElement element)
    {
        var cursor = element.NextSibling();
        while (cursor != null &&
               IsRangeMarker(cursor))
        {
            cursor = cursor.NextSibling();
        }

        if (cursor is Table table)
        {
            tablesToCheck.Add(table);
        }
    }

    /// <summary>
    /// Takes a block out of the document. Every removal of a tag paragraph, of a branch that was
    /// not chosen or of a loop body goes through here, since any of them can be all that kept two
    /// tables apart.
    /// </summary>
    void Remove(OpenXmlElement element)
    {
        RecordTableAfter(element);
        element.Remove();
    }

    /// <summary>
    /// A <c>w:sectPr</c> sits on the last paragraph of its section, so the paragraph a removed tag
    /// or token host leaves behind to carry one is an empty line at the end of that section. Run
    /// on the root runner once the whole tree has been processed: where a paragraph without a
    /// break of its own now precedes it, that paragraph ends the section instead and the stand-in
    /// goes. It stays after a table, which cannot carry a break, and wherever it is all the
    /// section has left.
    /// </summary>
    public void SettleSectionBreaks()
    {
        foreach (var carrier in sectionBreaks)
        {
            if (carrier.Parent == null ||
                PreviousBlock(carrier) is not Paragraph previous ||
                previous.ParagraphProperties?.SectionProperties != null)
            {
                continue;
            }

            var section = carrier.ParagraphProperties!.SectionProperties!;
            section.Remove();
            previous.ParagraphProperties ??= new();
            previous.ParagraphProperties.SectionProperties = section;
            carrier.Remove();
        }

        sectionBreaks.Clear();
    }

    void ApplyOne(StructuralReplacement replacement)
    {
        var host = replacement.Host;
        var parent = host.Parent;
        if (parent == null)
        {
            return;
        }

        if (replacement.KeepHost)
        {
            RecordTableAfter(InsertAfter(host, replacement.Produced));
            return;
        }

        // The token is all the text the host has, and what it produced takes the host's place.
        // Anything else the host carries is not the token's to take with it. A section break ends
        // the section after the produced content. Other content stays in a paragraph of its own,
        // before or after that content according to which side of the token it sat on.
        var section = host.ParagraphProperties?.SectionProperties;
        section?.Remove();

        Paragraph? tail = null;
        var keepHost = false;
        if (ParagraphContent.HasBesidesText(host))
        {
            // A field goes with the token, as it does when nothing else shares the host. The
            // split would otherwise leave its begin mark in one half and its end in the other.
            ParagraphContent.RemoveFieldMarks(host);
            var length = ParagraphText.Build(host).InnerText.Length;
            tail = (Paragraph)ParagraphSplicer.Split(host, 0, length, [], cloneIds)[0];
            // A page break before the host is ahead of everything that replaces it, so the host
            // half alone keeps it.
            tail.ParagraphProperties?.PageBreakBefore?.Remove();
            keepHost = ParagraphContent.HasBesidesText(host);
            if (!ParagraphContent.HasBesidesText(tail))
            {
                tail = null;
            }
        }

        var last = InsertAfter(host, replacement.Produced);
        if (tail != null)
        {
            last = parent.InsertAfter(tail, last);
        }

        if (section != null)
        {
            last = parent.InsertAfter(CarrySectionBreak(section), last);
        }

        // What follows now comes after the last of these or, when the token produced nothing and
        // the host kept nothing, after whatever preceded the host.
        RecordTableAfter(last);

        if (!keepHost)
        {
            host.Remove();
        }
    }

    OpenXmlElement InsertAfter(Paragraph host, IReadOnlyList<OpenXmlElement> elements)
    {
        var parent = host.Parent!;
        OpenXmlElement cursor = host;
        foreach (var produced in elements)
        {
            cursor = parent.InsertAfter(produced, cursor);
            if (produced is Table table)
            {
                tablesToCheck.Add(table);
            }
        }

        return cursor;
    }

    /// <summary>
    /// Takes a block tag paragraph out of the document. A tag paragraph is never part of the block
    /// it delimits, so what it carries besides the tag stays where the tag was, once: it is not
    /// repeated per iteration and does not go with an unchosen branch.
    /// </summary>
    void RemoveTag(Paragraph tag)
    {
        if (ParagraphContent.HasBesidesText(tag))
        {
            // Only the tag goes. A section break on the paragraph stays on it.
            ParagraphContent.RemoveText(tag);
            return;
        }

        if (tag.ParagraphProperties?.SectionProperties is { } section)
        {
            section.Remove();
            tag.InsertAfterSelf(CarrySectionBreak(section));
        }

        Remove(tag);
    }

    /// <summary>
    /// A tag paragraph that stays can anchor a text box, and the paragraphs in the box are token
    /// hosts of their own. They follow the tag in document order, so the scope tree lists them at
    /// the head of the block's body, where the runner for that body never finds them: the box
    /// stays with the tag, outside the block. They are run here instead, once, in the scope the
    /// tag paragraph itself is in.
    /// </summary>
    async Task ProcessCarriedAsync(Paragraph tag, IReadOnlyList<RangeNode> body)
    {
        if (!ParagraphContent.HasBesidesText(tag))
        {
            return;
        }

        foreach (var node in body)
        {
            if (!IsCarriedBy(tag, node))
            {
                return;
            }

            await ProcessAsync(node).ConfigureAwait(false);
        }
    }

    async Task ProcessCarriedAsync(IfNode ifNode, Paragraph open)
    {
        await ProcessCarriedAsync(open, ifNode.Branches[0].Body).ConfigureAwait(false);
        for (var i = 1; i < ifNode.Branches.Count; i++)
        {
            if (anchorMap.TryGetValue(ifNode.Branches[i].TagAnchorName, out var branchTag))
            {
                await ProcessCarriedAsync(branchTag, ifNode.Branches[i].Body).ConfigureAwait(false);
            }
        }

        if (ifNode.ElseAnchorName != null &&
            anchorMap.TryGetValue(ifNode.ElseAnchorName, out var elseTag))
        {
            await ProcessCarriedAsync(elseTag, ifNode.ElseBody).ConfigureAwait(false);
        }
    }

    bool IsCarriedBy(Paragraph tag, RangeNode node)
    {
        var anchor = node switch
        {
            SubstitutionNode substitution => substitution.AnchorName,
            StaticNode inert => inert.AnchorName,
            LoopNode loop => loop.OpenAnchorName,
            IfNode ifNode => ifNode.OpenAnchorName,
            _ => null
        };
        return anchor != null &&
               anchorMap.TryGetValue(anchor, out var paragraph) &&
               paragraph.Ancestors<Paragraph>().Contains(tag);
    }

    // An empty paragraph holding nothing but the break, where the paragraph that carried it was.
    Paragraph CarrySectionBreak(SectionProperties section)
    {
        var carrier = new Paragraph(
            new ParagraphProperties
            {
                SectionProperties = section
            });
        sectionBreaks.Add(carrier);
        return carrier;
    }

    async Task ProcessAsync(IReadOnlyList<RangeNode> nodes)
    {
        foreach (var node in nodes)
        {
            await ProcessAsync(node).ConfigureAwait(false);
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
        // lower-offset token. Structural ops (splice/split) do invalidate the cache — clear it
        // after those.
        ParagraphText? cachedText = null;
        ParagraphText Text() => cachedText ??= ParagraphText.Build(host);

        // Snapshot the host paragraph's original text once.
        var originalText = Text().InnerText;

        // node.Tokens come from TokenScan in ascending-offset order. Iterate in reverse without
        // allocating an OrderByDescending().ToList() — applying high-offset replacements first
        // keeps the lower-offset spans valid for the next iteration.
        List<(DocxTokenSite site, object value)>? soloStructuralTokens = null;
        var splitQueued = false;
        var tokenCount = node.Tokens.Count;
        // Only a structural-producing token needs to know whether an editable field shares the
        // paragraph, and that is the rare path — so defer the scan until one actually appears
        // rather than running HasEditableToken on every substitution node.
        bool? hasEditableSibling = null;
        // A MutateToken delegate can do anything to its paragraph, and every offset in node.Tokens
        // was measured before any of it. So only the token text is cleared in offset order; the
        // delegates run once nothing is left that depends on an offset.
        List<(MutateToken Mutate, Paragraph Target, DocxTokenSite Site)>? mutations = null;

        for (var i = tokenCount - 1; i >= 0; i--)
        {
            // Per token, so a render stops at the next token after the caller cancels rather than
            // at the end of the part. A token check is a volatile read: nothing is allocated.
            cancel.ThrowIfCancellationRequested();
            var token = node.Tokens[i];
            try
            {
                var evaluated = await EvaluateTokenAsync(token, host, tokenCount).ConfigureAwait(false);
                if (evaluated is EditableToken editable)
                {
                    if (editable.Entry.Kind == EditableFieldKind.Html)
                    {
                        QueueEditableHtml(host, token, editable, tokenCount, originalText);
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
                    hasEditableSibling ??= HasEditableToken(node);
                    if (hasEditableSibling.Value)
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

                    if (TokenScan.IsSolo(originalText, tokenCount, token.Offset, token.Length))
                    {
                        // The token is all the host paragraph says — queue for replacement after every
                        // other in-paragraph substitution has run.
                        soloStructuralTokens ??= [];
                        soloStructuralTokens.Add((token, evaluated));
                        continue;
                    }

                    var afterHalf = ApplyNonSoloStructural(token, host, evaluated, ref splitQueued);
                    cachedText = null;
                    if (afterHalf != null &&
                        mutations != null)
                    {
                        // The tokens already seen sit above the split, so their text went to the
                        // after half: that is the paragraph their delegates are owed.
                        for (var j = 0; j < mutations.Count; j++)
                        {
                            mutations[j] = (mutations[j].Mutate, afterHalf, mutations[j].Site);
                        }
                    }

                    continue;
                }

                if (evaluated is MutateToken mutate)
                {
                    // Clear the token text now, while its offset still holds. The paragraph is handed
                    // to the caller for in-place mutation after the loop.
                    Text().Replace(token.Offset, token.Length, string.Empty);
                    mutations ??= [];
                    mutations.Add((mutate, host, token));
                    continue;
                }

                var replacement = ToDisplayString(evaluated);
                Text().Replace(token.Offset, token.Length, replacement);
            }
            catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
            {
                // Everything a token does after it has a value: a block-shaped value building its
                // content, a MutateToken or OpenXmlToken delegate, an [ExcelsiorTable] build.
                throw WrapException(exception, host, token);
            }
        }

        if (mutations != null)
        {
            // In the order the tokens were reached: the last one in the paragraph first.
            foreach (var (mutate, target, site) in mutations)
            {
                var ctx = new OpenXmlContextImpl(mainPart, relationships.Owner, numberingState, styles.Value, target);
                // As for a produced fragment, see RenderTokenValue - but the target also holds what
                // the template put there, which refers to this part and is not to be repointed.
                var held = relationships.Snapshot();
                var before = held == null ? null : target.Descendants().ToHashSet();
                // The caller numbers a bookmark it adds without sight of the document, and the
                // paragraph it adds it to already holds the anchor's and maybe the template's -
                // see BookmarkIds.
                var existing = BookmarkIds.Markers(target);
                try
                {
                    mutate.Apply(target, ctx);
                }
                catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
                {
                    // Applied after the loop above, so outside its catch.
                    throw WrapException(exception, target, site);
                }
                relationships.Claim(target, before, held);
                bookmarks.RenumberAdded(target, existing);
            }
        }

        if (soloStructuralTokens is { Count: > 0 })
        {
            structuralReplacements ??= [];
            structuralReplacements.Add(new(host, BuildStructuralReplacements(host, soloStructuralTokens)));
        }
    }

    /// <summary>
    /// Queues an editable-HTML field as a block-level structural replacement. The member renders a
    /// rich-content control the user can edit within an editable range; it must occupy its own
    /// paragraph (like a read-only [Html] token) so the block <c>w:sdt</c> + perm range replace the
    /// host cleanly. Extraction serializes the block content back to HTML.
    /// </summary>
    void QueueEditableHtml(Paragraph host, DocxTokenSite token, EditableToken editable, int tokenCount, string originalText)
    {
        if (!TokenScan.IsSolo(originalText, tokenCount, token.Offset, token.Length))
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
            : numberingState.ConvertHtml(
                html,
                imagePolicies.BuildSettings(numberingSession: numberingState.GetHtmlSession()));
        bookmarks.Renumber(content);

        structuralReplacements ??= [];
        structuralReplacements.Add(new(host, EditableHtmlBuilder.Build(editable.Entry, content, editableState, host)));
    }

    /// <summary>
    /// Splice or split the host paragraph for a non-solo structural token. Inline-equivalent
    /// output (single produced paragraph) is unwrapped and spliced in place; anything else
    /// (multiple blocks, a table) splits the host paragraph at the token offset and inserts
    /// the produced elements between the two halves. Returns the after half when it split.
    /// </summary>
    Paragraph? ApplyNonSoloStructural(DocxTokenSite token, Paragraph host, object value, ref bool splitQueued)
    {
        var produced = RenderTokenValue(value, host, token, out var mayHoldSections);
        if (produced.Count == 0)
        {
            // Nothing to render — strip the token text from the host paragraph.
            ParagraphText.Build(host).Replace(token.Offset, token.Length, string.Empty);
            return null;
        }

        // A value that is nothing but a [SECTION] marker is a single paragraph too, but the break
        // is its paragraph properties, which are exactly what an inline splice drops. It takes the
        // split below instead, as any other block does.
        var loneMarker = mayHoldSections &&
                         produced.Count == 1 &&
                         SectionBreaks.IsMarker(produced[0]);
        if (!loneMarker &&
            ParagraphSplicer.IsInlineEquivalent(produced))
        {
            ParagraphSplicer.SpliceInline(host, token.Offset, token.Length, (Paragraph)produced[0], cloneIds);
            return null;
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
        // produced block elements between the resulting before/after halves. The host itself
        // becomes the before half, so the lower-offset tokens still to be applied keep running
        // against a paragraph that stays in the document; higher-offset tokens have already been
        // applied and are carried into the after half.
        var split = ParagraphSplicer.Split(host, token.Offset, token.Length, produced, cloneIds);
        if (mayHoldSections)
        {
            // Split hands back copies of what was produced, followed by the after half. The copies
            // are what reaches the document, and the after half is the host's own paragraph.
            SectionBreaks.Collect(split.Take(produced.Count), sectionMarkers);
        }

        structuralReplacements ??= [];
        structuralReplacements.Add(new(host, split, KeepHost: true));
        splitQueued = true;
        return (Paragraph)split[^1];
    }

    List<OpenXmlElement> BuildStructuralReplacements(Paragraph host, IReadOnlyList<(DocxTokenSite site, object value)> tokens)
    {
        var result = new List<OpenXmlElement>();
        foreach (var (site, value) in tokens)
        {
            try
            {
                var produced = RenderTokenValue(value, host, site, out var mayHoldSections);
                if (mayHoldSections)
                {
                    SectionBreaks.Collect(produced, sectionMarkers);
                }

                result.AddRange(produced);
            }
            catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
            {
                throw WrapException(exception, host, site);
            }
        }

        return result;
    }

    // mayHoldSections is true only for a markdown value, and only when its walk says so: a paragraph
    // carrying a w:sectPr is a [SECTION] marker there, where one an OpenXmlToken returns is a
    // section break the caller wrote whole.
    IReadOnlyList<OpenXmlElement> RenderTokenValue(object value, Paragraph host, DocxTokenSite site, out bool mayHoldSections)
    {
        // In a header, footer or note, what the html converter, Excelsior or a delegate adds to
        // the main part for this content belongs to the part the content lands in.
        var held = relationships.Snapshot();
        var produced = Produce(value, host, site, out mayHoldSections);
        relationships.Claim(produced, held);
        return produced;
    }

    IReadOnlyList<OpenXmlElement> Produce(object value, Paragraph host, DocxTokenSite site, out bool mayHoldSections)
    {
        if (value is MarkdownToken md)
        {
            try
            {
                // Renumbers its own bookmarks, from the same counter.
                return MarkdownRendering.Render(md.Source, mainPart, relationships.Owner, numberingState, imagePolicies, bookmarks, headingOffset: 0, out mayHoldSections, cancel);
            }
            catch (TableColumnLimitException exception)
            {
                // The table renderer knows neither the template nor the token its markdown came
                // through, so both are attached here.
                throw new ParchmentRenderException(
                    templateName,
                    exception.Message,
                    partUri,
                    Snippet(host, site),
                    site.Source);
            }
        }

        mayHoldSections = false;

        IReadOnlyList<OpenXmlElement> produced = value switch
        {
            HtmlToken html => numberingState.ConvertHtml(
                html.Source,
                imagePolicies.BuildSettings(numberingSession: numberingState.GetHtmlSession())),
            OpenXmlToken raw when ReferenceEquals(raw, OpenXmlToken.Empty) => [],
            OpenXmlToken raw => raw
                .Render(new OpenXmlContextImpl(mainPart, relationships.Owner, numberingState, styles.Value, host))
                .ToList(),
            _ => []
        };

        // Whatever produced these numbered its bookmarks without sight of the document they are
        // about to join - see BookmarkIds.
        bookmarks.Renumber(produced);
        return produced;
    }

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
        catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
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
        // text-token path doesn't need. A string assigned to a TokenValue member (a TextToken)
        // is on that path too: SharedFluid converts it, so it arrives here as a StringValue.
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
            return InterpretFluidValue(await pending.ConfigureAwait(false));
        }
        catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
        {
            throw WrapException(exception, host, site);
        }
    }

    ParchmentRenderException WrapException(Exception exception, Paragraph host, DocxTokenSite site) =>
        RenderFailure.Wrap(exception, templateName, partUri, Snippet(host, site), site.Source);

    // A block tag owns its paragraph, so the paragraph's text is the tag.
    ParchmentRenderException WrapTagException(Exception exception, Paragraph tag) =>
        RenderFailure.Wrap(exception, templateName, partUri, tokenSource: TagText(tag));

    static string TagText(Paragraph tag) =>
        ParagraphText.Build(tag).InnerText.Trim();

    OpenXmlToken? TryResolveExcelsiorTable(DocxTokenSite site)
    {
        if (excelsiorTables.IsEmpty ||
            site.References.Count == 0 ||
            LoopScope.Shadows(loopVariables, site))
        {
            return null;
        }

        // Match the token's full dotted reference (e.g. `Customer.Lines`) against the registered
        // map. Single-segment, multi-segment, and arbitrarily-nested paths from the root model
        // all flow through the same lookup. Loop-scope variables (e.g. `{{ line.SubItems }}`
        // inside `{% for line in Lines %}`) never match: the map is keyed on paths from the root
        // model only, and a loop variable that happens to spell a root path was ruled out above
        // — they fall through to normal Fluid evaluation. So does a token that indexes into the
        // member (`{{ Lines[0].Name }}`): its path keeps the index (IdentifierPath.Index) and
        // is no key of this map or of the three consulted after it.
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

        return new(_ => [ExcelsiorTableBridge.BuildTable(entry.ElementType, data, mainPart, bookmarks.Names, entry.HeadingParagraphStyle, entry.BodyParagraphStyle, entry.TableStyle, entry.Configure)]);
    }

    TokenValue? TryResolveStringList(DocxTokenSite site, Paragraph host, int siblingCount)
    {
        if (stringLists.IsEmpty ||
            site.References.Count == 0 ||
            LoopScope.Shadows(loopVariables, site))
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
        var paragraphText = ParagraphText.Build(host).InnerText;
        if (!TokenScan.IsSolo(paragraphText, siblingCount, site.Offset, site.Length))
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

    // No LoopScope check here, unlike the other three. The map a runner holds is never keyed
    // from the root while a loop is open: a loop body gets EditableMap.Empty, and an editable
    // collection's items get a map keyed under the loop variable itself (EditableMap.ScopedToItem),
    // which is exactly the token the check would turn away.
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
        if (formats.IsEmpty ||
            site.References.Count == 0 ||
            LoopScope.Shadows(loopVariables, site))
        {
            return null;
        }

        var reference = site.References[0];
        if (!formats.TryGet(reference.Dotted, out var entry))
        {
            return null;
        }

        // A marker on a non-string is a compile-time error (PARCH029), so the map holds strings
        // only. Should anything else arrive, it renders as an unmarked member would instead of as
        // nothing.
        var value = entry.Getter(rootModel);
        if (value is not (null or string))
        {
            return null;
        }

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

        try
        {
            await ProcessLoopAsync(loop, open, close, parent).ConfigureAwait(false);
        }
        catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
        {
            // The source, a modifier, or a lazily produced item. Whatever failed inside the body
            // was wrapped there, with its own token, and passes through.
            throw WrapTagException(exception, open);
        }
    }

    async Task ProcessLoopAsync(LoopNode loop, Paragraph open, Paragraph close, OpenXmlElement parent)
    {
        await ProcessCarriedAsync(open, loop.Body).ConfigureAwait(false);
        var items = await ResolveIterableAsync(loop, open).ConfigureAwait(false);
        var bodyElements = CaptureBetween(open, close);

        // A loop over an [EditableField] collection renders as an editable Word repeating section
        // rather than the read-only inline expansion. Body-only: header/footer parts get
        // EditableMap.Empty, so HasCollections is false there and the loop stays read-only.
        // loop.SourceReference is precomputed at tree-build time, so the common (no editable
        // collections) case costs a single bool check and never touches the AST.
        // A loop that holds no token for an editable member of the item has no control to put in a
        // row and nothing to extract, so it stays an ordinary read-only loop: the collection can be
        // listed again beside the loop that binds it (EditableTokenValidator allows one of those).
        if (editables.HasCollections &&
            loop.SourceReference is { } sourceReference &&
            editables.TryGetCollection(sourceReference, out var collectionEntry) &&
            HasItemControl(loop.Body, collectionEntry, loop.LoopVariable))
        {
            await RenderEditableCollection(loop, collectionEntry, open, parent, bodyElements, items.ToList()).ConfigureAwait(false);
            foreach (var element in bodyElements)
            {
                Remove(element);
            }

            RemoveTag(open);
            RemoveTag(close);
            return;
        }

        OpenXmlElement insertAnchor = open;
        var bodyIds = BodyIds(bodyElements);

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
            relationships,
            rootModel,
            excelsiorTables,
            formats,
            stringLists,
            EditableMap.Empty,
            editableState,
            bookmarks,
            bodyIds,
            numberingState,
            styles,
            imagePolicies,
            cancel,
            tablesToCheck,
            sectionMarkers,
            sectionBreaks)
        {
            loopVariables = [..loopVariables, loop.LoopVariable]
        };

        var first = true;
        foreach (var item in items)
        {
            cancel.ThrowIfCancellationRequested();

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
                    // The first copy keeps the ids the template gave the body; each later one
                    // needs its own - see CloneIds.
                    if (!first)
                    {
                        bodyIds?.Renumber(clone);
                    }

                    scratch.AppendChild(clone);
                    CollectAnchors(clone, cloneAnchors, loop.NestedAnchors);
                }

                first = false;

                // Reuse the original scope tree directly. cloneAnchors is keyed on the same anchor
                // names the registration-time tree references, so no per-iteration tree rebuild
                // (Remap) is needed. Multiple iterations leave duplicate-named bookmarks in the
                // live docx — that's fine because StripAll runs before Save and removes them all.
                await clonedRunner.RunAsync(loop.Body).ConfigureAwait(false);
                clonedRunner.ApplyStructural();

                // Snapshot scratch's children before mutating, so remove+insert doesn't break
                // the live ChildElements enumeration.
                pendingMoves.Clear();
                foreach (var element in scratch.ChildElements)
                {
                    pendingMoves.Add(element);
                }

                var iterationAnchor = insertAnchor;
                foreach (var produced in pendingMoves)
                {
                    produced.Remove();
                    insertAnchor = parent.InsertAfter(produced, insertAnchor);
                }

                // The block this iteration opens with now follows the previous iteration's last.
                RecordTableAfter(iterationAnchor);
            }
            finally
            {
                context.ReleaseScope();
            }
        }

        foreach (var element in bodyElements)
        {
            Remove(element);
        }

        RemoveTag(open);
        RemoveTag(close);
    }

    // Lockstep with EditableTokenValidator.Validate: the loop's own tokens, through any
    // conditional, but not those of a nested loop, which runs with an empty editable map.
    static bool HasItemControl(IReadOnlyList<RangeNode> nodes, CollectionEntry entry, string loopVariable)
    {
        foreach (var node in nodes)
        {
            switch (node)
            {
                case SubstitutionNode substitution:
                    foreach (var token in substitution.Tokens)
                    {
                        if (token.References.Count > 0 &&
                            entry.TryGetItemEntry(loopVariable, token.References[0].Dotted, out _))
                        {
                            return true;
                        }
                    }

                    break;

                case IfNode conditional:
                    foreach (var branch in conditional.Branches)
                    {
                        if (HasItemControl(branch.Body, entry, loopVariable))
                        {
                            return true;
                        }
                    }

                    if (HasItemControl(conditional.ElseBody, entry, loopVariable))
                    {
                        return true;
                    }

                    break;
            }
        }

        return false;
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
        // The blank item still binds the loop variable to a default element, so tokens that are not
        // editable fields resolve as they would for a real item. Its controls are built from no item
        // at all (see ScopedToItem): a default element's `0` or `false` would otherwise read back as
        // a value and the untouched row would extract as an item nobody entered.
        var isBlank = items.Count == 0;
        var effectiveItems = items.Count > 0
            ? items
            : [FluidValue.Create(entry.ElementFactory(), context.Options)];

        var scratch = new Body();
        var cloneAnchors = new Dictionary<string, Paragraph>(StringComparer.Ordinal);
        var repeatedItems = new List<SdtBlock>();
        IReadOnlyList<string> itemLoopVariables = [..loopVariables, loop.LoopVariable];
        var bodyIds = BodyIds(bodyElements);

        var first = true;
        foreach (var fluidItem in effectiveItems)
        {
            cancel.ThrowIfCancellationRequested();
            var rawItem = isBlank ? null : fluidItem.ToObjectValue();
            context.EnterChildScope();
            try
            {
                context.SetValue(loop.LoopVariable, fluidItem);
                cloneAnchors.Clear();
                foreach (var element in bodyElements)
                {
                    var clone = element.CloneNode(true);
                    // As in an ordinary loop: the template's ids go to the first item only.
                    if (!first)
                    {
                        bodyIds?.Renumber(clone);
                    }

                    scratch.AppendChild(clone);
                    CollectAnchors(clone, cloneAnchors, loop.NestedAnchors);
                }

                first = false;

                var itemRunner = new ScopeTreeRunner(
                    templateName,
                    partUri,
                    cloneAnchors,
                    context,
                    mainPart,
                    relationships,
                    rootModel,
                    excelsiorTables,
                    formats,
                    stringLists,
                    entry.ElementMap.ScopedToItem(loop.LoopVariable, rawItem),
                    editableState,
                    bookmarks,
                    bodyIds,
                    numberingState,
                    styles,
                    imagePolicies,
                    cancel,
                    tablesToCheck,
                    sectionMarkers,
                    sectionBreaks)
                {
                    loopVariables = itemLoopVariables
                };
                await itemRunner.RunAsync(loop.Body).ConfigureAwait(false);
                itemRunner.ApplyStructural();

                var itemBody = new List<OpenXmlElement>();
                foreach (var element in scratch.ChildElements.ToList())
                {
                    element.Remove();
                    itemBody.Add(element);
                }

                repeatedItems.Add(EditableRepeatingBuilder.BuildItem(itemBody, editableState, isBlank));
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

    /// <summary>
    /// The allocator the copies of a loop body are renumbered through, or null when the body holds
    /// no id a second copy would repeat. Asked once of the body rather than of every copy. What
    /// the body does not hold, nothing nested in it holds either, so the runner for its iterations
    /// is handed the same answer and its own loops and splits have nothing to walk.
    /// </summary>
    CloneIds? BodyIds(List<OpenXmlElement> bodyElements)
    {
        if (cloneIds != null &&
            CloneIds.Holds(bodyElements))
        {
            return cloneIds;
        }

        return null;
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
    static void CollectAnchors(OpenXmlElement clone, Dictionary<string, Paragraph> anchorMap, bool nestedAnchors)
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

            // A text box anchored in one of its runs holds paragraphs of its own, each with its
            // own anchor. Looking for them means walking the whole paragraph, so it is done only
            // for a block the scope tree marked as holding one — see LoopNode.NestedAnchors.
            if (!nestedAnchors)
            {
                return;
            }
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

    async Task<IEnumerable<FluidValue>> ResolveIterableAsync(LoopNode loop, Paragraph open)
    {
        // Hand the loop source straight to Fluid: ForStatement.Source is an Expression that, when
        // evaluated, yields a FluidValue we can enumerate. This honors filters, complex paths,
        // arithmetic, and any value converters Fluid is configured with — none of which the previous
        // reflection-walk supported.
        var sourceValue = await loop.LoopSource.EvaluateAsync(context).ConfigureAwait(false);
        var items = sourceValue.Enumerate(context);
        if (loop.Limit == null &&
            loop.Offset == null &&
            !loop.Reversed)
        {
            return items;
        }

        return await ApplyLoopModifiersAsync(loop, items.ToList(), open).ConfigureAwait(false);
    }

    // Same arithmetic as Fluid's ForStatement.WriteToAsync, which the markdown flow runs: offset
    // first, then limit, and `reversed` flips the window those two selected rather than the source,
    // whatever order the tag lists them in. A negative limit drops that many items from the end.
    async Task<IEnumerable<FluidValue>> ApplyLoopModifiersAsync(LoopNode loop, List<FluidValue> items, Paragraph open)
    {
        var start = 0;
        if (loop.Offset != null)
        {
            var offset = await loop.Offset.EvaluateAsync(context).ConfigureAwait(false);
            // Fluid indexes the source with a negative offset and throws. Nothing precedes the
            // first item, so the docx flow starts there instead.
            start = Math.Max(0, ModifierValue("offset", offset, open));
        }

        var count = Math.Max(0, items.Count - start);
        if (loop.Limit != null)
        {
            var limit = ModifierValue("limit", await loop.Limit.EvaluateAsync(context).ConfigureAwait(false), open);
            if (limit >= 0)
            {
                count = Math.Min(count, limit);
            }
            else
            {
                count = Math.Max(0, count + limit);
            }
        }

        if (count == 0)
        {
            return [];
        }

        var window = items.GetRange(start, count);
        if (loop.Reversed)
        {
            window.Reverse();
        }

        return window;
    }

    // The window is counted in ints, as the list it is cut from is. A value beyond that is
    // reported as what it is rather than left to overflow the cast.
    int ModifierValue(string modifier, FluidValue value, Paragraph open)
    {
        var number = value.ToNumberValue();
        if (number is >= int.MinValue and <= int.MaxValue)
        {
            return (int)number;
        }

        throw new ParchmentRenderException(
            templateName,
            $"The loop's '{modifier}' is {number.ToString(CultureInfo.InvariantCulture)}, which is outside the range a loop can be windowed by ({int.MinValue} to {int.MaxValue}).",
            partUri,
            tokenSource: TagText(open));
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
            var branch = ifNode.Branches[i];
            bool met;
            try
            {
                met = await EvaluateConditionAsync(branch.Condition).ConfigureAwait(false);
            }
            catch (Exception exception) when (RenderFailure.ShouldWrap(exception))
            {
                // The paragraph of the branch's own tag, so an {% elsif %} is named as itself.
                throw WrapTagException(exception, anchorMap.GetValueOrDefault(branch.TagAnchorName, open));
            }

            if (!met)
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
        var branchTags = BranchTags(ifNode);
        await ProcessCarriedAsync(ifNode, open).ConfigureAwait(false);

        if (chosen == null)
        {
            foreach (var element in allBranchElements)
            {
                RemoveBranchElement(element, branchTags);
            }

            RemoveTag(open);
            RemoveTag(close);
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
            CollectAnchors(element, branchAnchors, ifNode.NestedAnchors);
        }

        var innerRunner = new ScopeTreeRunner(templateName, partUri, branchAnchors, context, mainPart, relationships, rootModel, excelsiorTables, formats, stringLists, editables, editableState, bookmarks, cloneIds, numberingState, styles, imagePolicies, cancel, tablesToCheck, sectionMarkers, sectionBreaks)
        {
            loopVariables = loopVariables
        };
        await innerRunner.RunAsync(chosen).ConfigureAwait(false);
        innerRunner.ApplyStructural();

        var keep = new HashSet<OpenXmlElement>(keepElements);
        foreach (var element in allBranchElements)
        {
            if (!keep.Contains(element))
            {
                RemoveBranchElement(element, branchTags);
            }
        }

        RemoveTag(open);
        RemoveTag(close);
    }

    /// <summary>
    /// The <c>{% elsif %}</c> and <c>{% else %}</c> tag paragraphs of <paramref name="ifNode"/>,
    /// or null when it has neither. They sit among the branch content but belong to no branch.
    /// </summary>
    List<Paragraph>? BranchTags(IfNode ifNode)
    {
        List<Paragraph>? tags = null;
        for (var i = 1; i < ifNode.Branches.Count; i++)
        {
            if (anchorMap.TryGetValue(ifNode.Branches[i].TagAnchorName, out var branchTag))
            {
                tags ??= [];
                tags.Add(branchTag);
            }
        }

        if (ifNode.ElseAnchorName != null &&
            anchorMap.TryGetValue(ifNode.ElseAnchorName, out var elseTag))
        {
            tags ??= [];
            tags.Add(elseTag);
        }

        return tags;
    }

    void RemoveBranchElement(OpenXmlElement element, List<Paragraph>? branchTags)
    {
        if (branchTags != null &&
            element is Paragraph paragraph &&
            branchTags.Contains(paragraph))
        {
            RemoveTag(paragraph);
            return;
        }

        Remove(element);
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
        (await condition.EvaluateAsync(context).ConfigureAwait(false)).ToBooleanValue();

    static string ToDisplayString(object value) =>
        value switch
        {
            null => string.Empty,
            string s => s,
            _ => value.ToString() ?? string.Empty
        };

    static string Snippet(Paragraph host, DocxTokenSite site)
    {
        // Clamped at both ends: a failure can be reported after the token's text has been replaced
        // or its paragraph mutated, when the site's offsets no longer describe the paragraph.
        var text = ParagraphText.Build(host).InnerText;
        var start = Math.Clamp(site.Offset - 40, 0, text.Length);
        var end = Math.Clamp(site.Offset + site.Length + 40, start, text.Length);
        return text[start..end];
    }

    sealed record StructuralReplacement(Paragraph Host, IReadOnlyList<OpenXmlElement> Produced, bool KeepHost = false);
}
