/// <summary>
/// Registration-time validator for substitution tokens that resolve to an
/// <see cref="ExcelsiorTableAttribute"/>-marked property. Such tokens become structural
/// replacements at render time (the entire host paragraph is swapped for the rendered table),
/// which means two things the user must get right up front:
/// <list type="number">
///   <item>The token has to sit alone in its own paragraph — any surrounding static text or
///     sibling tokens would be silently discarded by the paragraph swap.</item>
///   <item>The token has to be a plain identifier path (e.g. <c>{{ Buyer.Addresses }}</c>).
///     Filters, arithmetic, and literals are rejected because we bypass Fluid evaluation for
///     Excelsior tokens and walk the model object directly, so filter output would never be
///     applied.</item>
/// </list>
/// Both conditions are caught at registration so template authors see the error immediately
/// instead of debugging a surprising rendered document.
/// </summary>
static class ExcelsiorTokenValidator
{
    public static void Validate(
        IReadOnlyList<ParagraphClassification> classifications,
        ExcelsiorTableMap excelsiorTables,
        string templateName,
        string partUri)
    {
        if (excelsiorTables.IsEmpty)
        {
            return;
        }

        foreach (var classification in classifications)
        {
            if (classification.Kind != ParagraphKind.Substitution)
            {
                continue;
            }

            var paragraphText = ParagraphText.Build(classification.Paragraph).InnerText;
            foreach (var token in classification.Substitutions)
            {
                if (token.References.Count == 0)
                {
                    continue;
                }

                if (!excelsiorTables.TryGet(token.References[0].Dotted, out _))
                {
                    continue;
                }

                RequireSoloInParagraph(
                    token,
                    classification.Substitutions.Count,
                    paragraphText,
                    templateName,
                    partUri);

                RequirePlainIdentifier(token, templateName, partUri);
            }
        }
    }

    static void RequireSoloInParagraph(
        DocxTokenSite token,
        int siblingTokenCount,
        string paragraphText,
        string templateName,
        string partUri)
    {
        if (siblingTokenCount == 1 &&
            token.Offset == 0 &&
            token.Length == paragraphText.Length)
        {
            return;
        }

        throw new ParchmentRegistrationException(
            templateName,
            $"""
             [ExcelsiorTable] token '{token.Source}' must sit alone in its own paragraph.
             Structural table replacement swaps the entire host paragraph, so any surrounding text or sibling tokens would be discarded.
             """,
            partUri,
            token.Source);
    }

    static void RequirePlainIdentifier(DocxTokenSite token, string templateName, string partUri)
    {
        var statements = ((Fluid.Parser.FluidTemplate) token.Template).Statements;
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
               [ExcelsiorTable] token '{{{token.Source}}}' must be a plain member-access expression (for example '{{ Lines }}' or '{{ Buyer.Addresses }}').
               Filters, arithmetic, and literal expressions are not supported — the Excelsior rendering path walks the model object directly and bypasses Fluid evaluation.
               """,
            partUri,
            token.Source);
    }
}
