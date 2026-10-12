/// <summary>
/// Registration-time validator for substitution tokens that resolve to an
/// <c>[EditableField]</c> member. Rules:
/// <list type="bullet">
/// <item>plain member-access expression only — the editable render path bypasses Fluid, so a
/// filter chain would be silently ignored;</item>
/// <item>each editable path may appear at most once in the document body — the dotted path is
/// the content control's tag, and extraction needs tags to be unique;</item>
/// <item>an editable token may not share its paragraph with a structural token
/// (<c>[ExcelsiorTable]</c> / <c>[Html]</c> / <c>[Markdown]</c>) — structural splice/split
/// machinery clones paragraph halves, which would corrupt the field's perm-range markers.</item>
/// <item>a body loop over an <c>[EditableField]</c> collection may not carry <c>limit:</c> /
/// <c>offset:</c> / <c>reversed</c> — extraction replaces the whole collection with the rows the
/// document holds, so a windowed or reordered render would write that subset or order back.</item>
/// <item>an <c>[EditableField]</c> collection may be bound once: one loop holds the tokens for the
/// item's editable members, each of them once — extraction reads one repeating section per
/// collection and one control per member in a row, so the edits made in a second would be
/// discarded. A loop that holds none of them renders read-only and is not a binding.</item>
/// </list>
/// </summary>
static class EditableTokenValidator
{
    public static void Validate(
        IReadOnlyList<ParagraphClassification> classifications,
        EditableMap editables,
        ExcelsiorTableMap excelsiorTables,
        FormatMap formats,
        string templateName,
        string partUri,
        bool isBody)
    {
        if (editables.IsEmpty)
        {
            return;
        }

        var seen = isBody ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;

        // The top-level body loop being walked, when its source is an editable collection, with
        // the item members its tokens have bound so far.
        CollectionEntry? collectionLoop = null;
        var boundMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var boundCollections = isBody ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;

        var loopVariables = new List<string>();
        foreach (var classification in classifications)
        {
            // Only a top-level body loop renders as a repeating section: non-body parts and
            // nested loop bodies run with an empty editable map and stay read-only, where a
            // modifier is harmless. Checked before the tag opens its own variable.
            if (isBody &&
                loopVariables.Count == 0 &&
                classification.Block is { Kind: BlockTagKind.For } block)
            {
                RejectModifiersOnCollectionLoop(block, editables, templateName, partUri);
                collectionLoop = CollectionOf(block, editables);
                boundMembers.Clear();
            }

            LoopScope.Track(loopVariables, classification);
            if (loopVariables.Count == 0)
            {
                collectionLoop = null;
            }

            if (classification.Kind != ParagraphKind.Substitution)
            {
                continue;
            }

            // Only the loop's own tokens become controls: a loop nested in it runs with an empty
            // editable map, the same as any other loop body.
            if (boundCollections != null &&
                collectionLoop != null &&
                loopVariables.Count == 1)
            {
                RejectSecondBinding(classification, collectionLoop, loopVariables[0], boundMembers, boundCollections, templateName, partUri);
            }

            List<DocxTokenSite>? editableTokens = null;
            foreach (var token in classification.Substitutions)
            {
                // A token that starts with a loop variable reads the loop item, whatever root
                // path it happens to spell — it is not a use of the root's editable member.
                if (token.References.Count == 0 ||
                    LoopScope.Shadows(loopVariables, token))
                {
                    continue;
                }

                if (!editables.TryGet(token.References[0].Dotted, out _))
                {
                    continue;
                }

                editableTokens ??= [];
                editableTokens.Add(token);
            }

            if (editableTokens == null)
            {
                continue;
            }

            foreach (var token in editableTokens)
            {
                RequirePlainIdentifier(token, templateName, partUri);

                if (seen != null &&
                    !seen.Add(token.References[0].Dotted))
                {
                    throw new ParchmentRegistrationException(
                        templateName,
                        $"[EditableField] token '{token.Source}' appears more than once in the document body. The dotted path is the content control's tag and must be unique for extraction — reference each editable member once.",
                        partUri,
                        token.Source);
                }
            }

            if (classification.Substitutions.Count == editableTokens.Count)
            {
                continue;
            }

            foreach (var token in classification.Substitutions)
            {
                if (token.References.Count == 0 ||
                    LoopScope.Shadows(loopVariables, token))
                {
                    continue;
                }

                var dotted = token.References[0].Dotted;
                if (excelsiorTables.TryGet(dotted, out _) ||
                    formats.TryGet(dotted, out _))
                {
                    throw new ParchmentRegistrationException(
                        templateName,
                        $"Token '{token.Source}' produces structural content but shares its paragraph with an editable field ('{editableTokens[0].Source}'). Structural replacement clones paragraph halves and would corrupt the field's editable range. Move one of the tokens to its own paragraph.",
                        partUri,
                        token.Source);
                }
            }
        }
    }

    static void RejectModifiersOnCollectionLoop(BlockMarker block, EditableMap editables, string templateName, string partUri)
    {
        if (block.LoopLimit == null &&
            block.LoopOffset == null &&
            !block.LoopReversed)
        {
            return;
        }

        var collection = CollectionOf(block, editables);
        if (collection == null)
        {
            return;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $"Loop '{block.Source}' iterates the [EditableField] collection '{collection.DottedPath}' with a limit, offset or reversed modifier. The loop renders as a repeating section and extraction replaces the whole collection with the rows the document holds, so a windowed or reordered render would be written back as the collection. Remove the modifier, or drop [EditableField] from the member to render it read-only.",
            partUri,
            block.Source);
    }

    static CollectionEntry? CollectionOf(BlockMarker block, EditableMap editables)
    {
        var sourceRefs = IdentifierVisitor.Collect(block.LoopSource!);
        if (sourceRefs.Count == 0 ||
            !editables.TryGetCollection(sourceRefs[0].Dotted, out var collection))
        {
            return null;
        }

        return collection;
    }

    static void RejectSecondBinding(
        ParagraphClassification classification,
        CollectionEntry collection,
        string loopVariable,
        HashSet<string> boundMembers,
        HashSet<string> boundCollections,
        string templateName,
        string partUri)
    {
        foreach (var token in classification.Substitutions)
        {
            if (token.References.Count == 0 ||
                !collection.TryGetItemEntry(loopVariable, token.References[0].Dotted, out var entry))
            {
                continue;
            }

            if (!boundMembers.Add(entry.DottedPath))
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] token '{token.Source}' appears more than once in the loop over '{collection.DottedPath}'. The member path is the content control's tag within a row and must be unique for extraction — reference each editable member of the item once.",
                    partUri,
                    token.Source);
            }

            // The first control of a loop is what makes it a binding of the collection.
            if (boundMembers.Count == 1 &&
                !boundCollections.Add(collection.DottedPath))
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] token '{token.Source}' binds the collection '{collection.DottedPath}' in a second loop, so the collection's editable members appear more than once in the document body. Each such loop renders as a repeating section and extraction reads only one — keep the item's editable members in one loop. A loop that prints only members that are not [EditableField] renders read-only and is fine.",
                    partUri,
                    token.Source);
            }
        }
    }

    static void RequirePlainIdentifier(DocxTokenSite token, string templateName, string partUri)
    {
        var statements = ((Fluid.Parser.FluidTemplate)token.Template).Statements;
        if (statements.Count == 0)
        {
            return;
        }

        if (statements[0] is not OutputStatement output)
        {
            return;
        }

        if (output.Expression is MemberExpression)
        {
            return;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $$$"""
               [EditableField] token '{{{token.Source}}}' must be a plain member-access expression (for example '{{ PurchaseOrder }}' or '{{ Customer.Notes }}').
               Filters, arithmetic, and literal expressions are not supported — the editable render path is selected by attribute, so filters would not be applied.
               """,
            partUri,
            token.Source);
    }
}
