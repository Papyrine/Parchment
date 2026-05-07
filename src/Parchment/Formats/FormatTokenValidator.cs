/// <summary>
/// Registration-time validator for substitution tokens that resolve to a string property
/// annotated as html or markdown. The remaining rule (after relaxing the solo-in-paragraph
/// requirement in favor of inline splicing / paragraph splitting at render time) is:
/// the token must be a plain member-access expression (no filters / arithmetic / literals),
/// because the formatted rendering is selected by attribute and a filter chain would be silently
/// ignored otherwise.
/// </summary>
static class FormatTokenValidator
{
    public static void Validate(
        IReadOnlyList<ParagraphClassification> classifications,
        FormatMap formats,
        string templateName,
        string partUri)
    {
        if (formats.IsEmpty)
        {
            return;
        }

        foreach (var classification in classifications)
        {
            if (classification.Kind != ParagraphKind.Substitution)
            {
                continue;
            }

            foreach (var token in classification.Substitutions)
            {
                if (token.References.Count == 0)
                {
                    continue;
                }

                if (!formats.TryGet(token.References[0].Dotted, out var kind))
                {
                    continue;
                }

                RequirePlainIdentifier(token, kind, templateName, partUri);
            }
        }
    }

    static void RequirePlainIdentifier(DocxTokenSite token, FormatKind kind, string templateName, string partUri)
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
               [{{{kind}}}] token '{{{token.Source}}}' must be a plain member-access expression (for example '{{ Body }}' or '{{ Customer.Bio }}').
               Filters, arithmetic, and literal expressions are not supported — the property's formatted rendering is selected by attribute, so filters would not be applied.
               """,
            partUri,
            token.Source);
    }
}
