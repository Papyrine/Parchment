/// <summary>
/// Registration-time validator for substitution tokens that resolve to an
/// <c>[EditableField]</c> member. The rules apply wherever the runner renders the token as a
/// control — a root member outside every loop, and a member of an <c>[EditableField]</c>
/// collection's item reached through the loop variable directly inside that loop — and nowhere
/// else, since every other occurrence renders as plain read-only text. Rules:
/// <list type="bullet">
/// <item>plain member-access expression only — the editable render path bypasses Fluid, so a
/// filter chain would be silently ignored;</item>
/// <item>each editable path may appear at most once in the document body, and each item member
/// at most once in its collection's loop — the dotted path is the content control's tag, and
/// extraction needs tags to be unique. The <c>mc:Fallback</c> copy of a text box is not an
/// appearance: Word stores a text box twice, the fallback mirrors the <c>mc:Choice</c> copy, and
/// extraction reads the choice alone;</item>
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
/// <item>a token that renders as a content control may not sit inside a content control of the
/// author's that holds a single typed value (plain text, date picker, drop-down list, combo box,
/// checkbox, picture) — <see cref="EditableSplicer"/> puts the field where the token is, and Word
/// refuses to open a document with a control nested in one of those. A rich-text control is
/// fine.</item>
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
        // Only the body renders controls: every other part runs with an empty editable map, so a
        // token there is an ordinary read-only substitution and none of the rules apply to it.
        if (editables.IsEmpty ||
            !isBody)
        {
            return;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // The item map and the paths bound so far while a top-level loop over an [EditableField]
        // collection is open. The map is the projection the runner renders each item with, so a
        // token is matched to an item member here exactly when it becomes a control there.
        EditableMap? itemEditables = null;
        HashSet<string>? itemSeen = null;

        // The collection whose top-level loop is open, and the collections an earlier loop has
        // already bound.
        CollectionEntry? collectionLoop = null;
        var boundCollections = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var loopVariables = new List<string>();
        foreach (var classification in classifications)
        {
            // Only a top-level body loop renders as a repeating section: a nested loop body runs
            // with an empty editable map and stays read-only, where a modifier is harmless.
            // Checked before the tag opens its own variable.
            if (loopVariables.Count == 0 &&
                classification.Block is { Kind: BlockTagKind.For } block &&
                TryGetCollection(block, editables, out var collection))
            {
                RejectModifiersOnCollectionLoop(block, collection, templateName, partUri);
                itemEditables = collection.ElementMap.ScopedToItem(block.LoopVariable!, null);
                itemSeen = new(StringComparer.OrdinalIgnoreCase);
                collectionLoop = collection;
            }

            LoopScope.Track(loopVariables, classification);
            if (loopVariables.Count == 0)
            {
                itemEditables = null;
                itemSeen = null;
                collectionLoop = null;
            }

            if (classification.Kind != ParagraphKind.Substitution)
            {
                continue;
            }

            // Which map the runner holds here: the root's outside every loop, the item's directly
            // inside a repeating section, and none anywhere deeper — a loop body renders editable
            // members as plain text, so a filter or a second occurrence there is harmless.
            var active = editables;
            var activeSeen = seen;
            if (loopVariables.Count > 0)
            {
                if (loopVariables.Count > 1 ||
                    itemEditables == null)
                {
                    continue;
                }

                active = itemEditables;
                activeSeen = itemSeen!;
            }

            List<DocxTokenSite>? editableTokens = null;
            foreach (var token in classification.Substitutions)
            {
                // No LoopScope check, as in ScopeTreeRunner.TryResolveEditable: the root map is
                // only consulted with no loop open, and the item map is keyed under the loop
                // variable itself.
                if (token.References.Count == 0 ||
                    !active.TryGet(token.References[0].Dotted, out _))
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

            // A text box is stored as a DrawingML shape in mc:Choice and again as VML in
            // mc:Fallback, so a token typed once into the box arrives here twice. The fallback
            // copy still renders a control, to stay the mirror of the choice, but extraction
            // skips it (ParchmentExtractor.InFallback), so it cannot make a tag ambiguous.
            var counts = !InFallback(classification.Paragraph);

            foreach (var token in editableTokens)
            {
                RejectInsideValueControl(classification.Paragraph, token, templateName, partUri);
                RequirePlainIdentifier(token, templateName, partUri);

                if (!counts)
                {
                    continue;
                }

                if (activeSeen.Add(token.References[0].Dotted))
                {
                    // The first control of a loop is what makes it a binding of the collection.
                    if (collectionLoop != null &&
                        loopVariables.Count > 0 &&
                        activeSeen.Count == 1 &&
                        !boundCollections.Add(collectionLoop.DottedPath))
                    {
                        throw new ParchmentRegistrationException(
                            templateName,
                            $"[EditableField] token '{token.Source}' binds the collection '{collectionLoop.DottedPath}' in a second loop, so the collection's editable members appear more than once in the document body. Each such loop renders as a repeating section and extraction reads only one — keep the item's editable members in one loop. A loop that prints only members that are not [EditableField] renders read-only and is fine.",
                            partUri,
                            token.Source);
                    }

                    continue;
                }

                if (loopVariables.Count > 0)
                {
                    throw new ParchmentRegistrationException(
                        templateName,
                        $"[EditableField] token '{token.Source}' appears more than once in the loop over its [EditableField] collection. The member's path is the content control's tag inside each repeated item and must be unique there for extraction — reference each editable member once.",
                        partUri,
                        token.Source);
                }

                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] token '{token.Source}' appears more than once in the document body. The dotted path is the content control's tag and must be unique for extraction — reference each editable member once.",
                    partUri,
                    token.Source);
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

    static bool InFallback(Paragraph paragraph)
    {
        foreach (var ancestor in paragraph.Ancestors())
        {
            if (ancestor is AlternateContentFallback)
            {
                return true;
            }
        }

        return false;
    }

    // Asked only of a token that becomes a content control: a root member outside every loop, or
    // an item member directly inside the top-level loop over an editable collection. Anything
    // else inside a loop renders as read-only text, which any control can hold.
    static void RejectInsideValueControl(Paragraph paragraph, DocxTokenSite token, string templateName, string partUri)
    {
        var kind = ValueControlAround(paragraph, token.Offset);
        if (kind == null)
        {
            return;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $"[EditableField] token '{token.Source}' sits inside a {kind} content control. The token renders as a content control of its own, and Word refuses to open a document with a control nested in one that holds a single value. Move the token out of the control, or change the control to rich text.",
            partUri,
            token.Source);
    }

    /// <summary>
    /// The kind of the nearest single-value content control around the character at
    /// <paramref name="offset"/> — where the field is inserted — or null when there is none.
    /// </summary>
    static string? ValueControlAround(Paragraph paragraph, int offset)
    {
        var consumed = 0;
        foreach (var text in ParagraphText.Texts(paragraph))
        {
            consumed += text.Text.Length;
            if (offset >= consumed)
            {
                continue;
            }

            foreach (var sdt in text.Ancestors<SdtElement>())
            {
                var kind = ValueControlKind(sdt);
                if (kind != null)
                {
                    return kind;
                }
            }

            return null;
        }

        return null;
    }

    // Lockstep with the source generator's DocxArchiveReader.IsValueControl.
    static string? ValueControlKind(SdtElement sdt)
    {
        var properties = sdt.GetFirstChild<SdtProperties>();
        if (properties == null)
        {
            return null;
        }

        foreach (var child in properties.ChildElements)
        {
            switch (child)
            {
                case SdtContentText:
                    return "plain-text";
                case SdtContentDate:
                    return "date picker";
                case SdtContentDropDownList:
                    return "drop-down list";
                case SdtContentComboBox:
                    return "combo box";
                case DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox:
                    return "checkbox";
                case SdtContentPicture:
                    return "picture";
            }
        }

        return null;
    }

    static bool TryGetCollection(BlockMarker block, EditableMap editables, [NotNullWhen(true)] out CollectionEntry? collection)
    {
        collection = null;
        var sourceRefs = IdentifierVisitor.Collect(block.LoopSource!);
        return sourceRefs.Count > 0 &&
               editables.TryGetCollection(sourceRefs[0].Dotted, out collection);
    }

    static void RejectModifiersOnCollectionLoop(BlockMarker block, CollectionEntry collection, string templateName, string partUri)
    {
        if (block.LoopLimit == null &&
            block.LoopOffset == null &&
            !block.LoopReversed)
        {
            return;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $"Loop '{block.Source}' iterates the [EditableField] collection '{collection.DottedPath}' with a limit, offset or reversed modifier. The loop renders as a repeating section and extraction replaces the whole collection with the rows the document holds, so a windowed or reordered render would be written back as the collection. Remove the modifier, or drop [EditableField] from the member to render it read-only.",
            partUri,
            block.Source);
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
