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
        if (editables.IsEmpty)
        {
            return;
        }

        var seen = isBody ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;

        var loopVariables = new List<string>();
        CollectionEntry? collection = null;
        foreach (var classification in classifications)
        {
            if (isBody)
            {
                collection = TrackCollection(collection, loopVariables, classification, editables);
            }

            // Only a top-level body loop renders as a repeating section: non-body parts and
            // nested loop bodies run with an empty editable map and stay read-only, where a
            // modifier is harmless. Checked before the tag opens its own variable.
            if (isBody &&
                loopVariables.Count == 0 &&
                classification.Block is { Kind: BlockTagKind.For } block)
            {
                RejectModifiersOnCollectionLoop(block, editables, templateName, partUri);
            }

            LoopScope.Track(loopVariables, classification);
            if (classification.Kind != ParagraphKind.Substitution)
            {
                continue;
            }

            if (isBody)
            {
                RejectInsideValueControl(classification, loopVariables, collection, editables, templateName, partUri);
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

    /// <summary>
    /// The editable collection whose top-level loop the walk is inside, if any. Called before
    /// <see cref="LoopScope.Track"/> moves <paramref name="loopVariables"/> past the paragraph.
    /// </summary>
    static CollectionEntry? TrackCollection(
        CollectionEntry? current,
        List<string> loopVariables,
        ParagraphClassification classification,
        EditableMap editables)
    {
        switch (classification.Block?.Kind)
        {
            case BlockTagKind.For when loopVariables.Count == 0:
                var sourceRefs = IdentifierVisitor.Collect(classification.Block.LoopSource!);
                if (sourceRefs.Count > 0 &&
                    editables.TryGetCollection(sourceRefs[0].Dotted, out var entry))
                {
                    return entry;
                }

                return null;
            case BlockTagKind.EndFor when loopVariables.Count == 1:
                return null;
        }

        return current;
    }

    // Only the tokens that become content controls: a root member outside every loop, or an item
    // member directly inside the top-level loop over an editable collection. Anything else inside
    // a loop renders as read-only text, which any control can hold.
    static void RejectInsideValueControl(
        ParagraphClassification classification,
        List<string> loopVariables,
        CollectionEntry? collection,
        EditableMap editables,
        string templateName,
        string partUri)
    {
        foreach (var token in classification.Substitutions)
        {
            if (token.References.Count == 0 ||
                !RendersAsControl(token.References[0], loopVariables, collection, editables))
            {
                continue;
            }

            var kind = ValueControlAround(classification.Paragraph, token.Offset);
            if (kind == null)
            {
                continue;
            }

            throw new ParchmentRegistrationException(
                templateName,
                $"[EditableField] token '{token.Source}' sits inside a {kind} content control. The token renders as a content control of its own, and Word refuses to open a document with a control nested in one that holds a single value. Move the token out of the control, or change the control to rich text.",
                partUri,
                token.Source);
        }
    }

    static bool RendersAsControl(
        IdentifierPath path,
        List<string> loopVariables,
        CollectionEntry? collection,
        EditableMap editables)
    {
        if (loopVariables.Count == 0)
        {
            return editables.TryGet(path.Dotted, out _);
        }

        if (collection == null ||
            loopVariables.Count != 1 ||
            path.Segments.Count < 2 ||
            !string.Equals(path.Root, loopVariables[0], StringComparison.Ordinal))
        {
            return false;
        }

        return collection.ElementMap.TryGet(string.Join('.', path.Segments.Skip(1)), out _);
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

    static void RejectModifiersOnCollectionLoop(BlockMarker block, EditableMap editables, string templateName, string partUri)
    {
        if (block.LoopLimit == null &&
            block.LoopOffset == null &&
            !block.LoopReversed)
        {
            return;
        }

        var sourceRefs = IdentifierVisitor.Collect(block.LoopSource!);
        if (sourceRefs.Count == 0 ||
            !editables.TryGetCollection(sourceRefs[0].Dotted, out _))
        {
            return;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $"Loop '{block.Source}' iterates the [EditableField] collection '{sourceRefs[0].Dotted}' with a limit, offset or reversed modifier. The loop renders as a repeating section and extraction replaces the whole collection with the rows the document holds, so a windowed or reordered render would be written back as the collection. Remove the modifier, or drop [EditableField] from the member to render it read-only.",
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
