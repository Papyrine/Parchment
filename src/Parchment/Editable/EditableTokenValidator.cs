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

        foreach (var classification in classifications)
        {
            if (classification.Kind != ParagraphKind.Substitution)
            {
                continue;
            }

            List<DocxTokenSite>? editableTokens = null;
            foreach (var token in classification.Substitutions)
            {
                if (token.References.Count == 0)
                {
                    continue;
                }

                if (!editables.TryGet(token.References[0].Dotted, out _))
                {
                    continue;
                }

                (editableTokens ??= []).Add(token);
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
                if (token.References.Count == 0)
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
