/// <summary>
/// Scans every paragraph of a given part for Fluid tokens, classifies each paragraph, and injects
/// anchor bookmarks into token-bearing paragraphs so they can be located by name at render time.
/// </summary>
static class TokenScanner
{
    public static List<ParagraphClassification> Scan(OpenXmlCompositeElement partRoot, string templateName, string partUri)
    {
        var results = new List<ParagraphClassification>();
        // Seed the anchor id counter once for the whole part, then hand out incrementing ids as
        // anchors are inserted — see Anchors.MaxBookmarkId / EnsureOn.
        var nextBookmarkId = Anchors.MaxBookmarkId(partRoot) + 1;
        foreach (var paragraph in partRoot.Descendants<Paragraph>().ToList())
        {
            var classification = ClassifyParagraph(paragraph, templateName, partUri);
            if (classification.Kind == ParagraphKind.Static)
            {
                continue;
            }

            var anchor = Anchors.EnsureOn(paragraph, ref nextBookmarkId);
            results.Add(classification with { AnchorName = anchor });
        }

        return results;
    }

    static ParagraphClassification ClassifyParagraph(Paragraph paragraph, string templateName, string partUri)
    {
        var text = ParagraphText.Build(paragraph);
        var innerText = text.InnerText;
        var sites = TokenScan.Scan(innerText);
        if (sites.Count == 0)
        {
            return new(paragraph, string.Empty, ParagraphKind.Static, [], null);
        }

        var substitutions = new List<DocxTokenSite>();
        var blocks = new List<BlockMarker>();

        foreach (var site in sites)
        {
            var source = innerText.Substring(site.Offset, site.Length);
            if (site.Kind == TokenSiteKind.Substitution)
            {
                substitutions.Add(ParseSubstitution(source, site.Offset, templateName, partUri));
            }
            else
            {
                blocks.Add(ParseBlockTag(source, templateName, partUri));
            }
        }

        if (blocks.Count > 1)
        {
            throw new ParchmentRegistrationException(
                templateName,
                "A paragraph contains more than one block tag. Place each block tag in its own paragraph.",
                partUri,
                string.Join(' ', blocks.Select(x => x.Source)));
        }

        if (blocks.Count == 1)
        {
            if (substitutions.Count > 0)
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    "A paragraph mixes block tags with substitution tokens. Block tags must sit in their own paragraphs.",
                    partUri,
                    blocks[0].Source);
            }

            if (TokenScan.HasContentOutsideSites(innerText, sites))
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    "A paragraph contains a block tag alongside non-token text. Block tags must sit in their own paragraphs.",
                    partUri,
                    blocks[0].Source);
            }

            return new(paragraph, string.Empty, ParagraphKind.Block, [], blocks[0]);
        }

        return new(paragraph, string.Empty, ParagraphKind.Substitution, substitutions, null);
    }

    static DocxTokenSite ParseSubstitution(string source, int offset, string templateName, string partUri)
    {
        if (!SharedFluid.Parser.TryParse(source, out var template, out var error))
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Failed to parse liquid token: {error}",
                partUri,
                source);
        }

        var refs = IdentifierVisitor.Collect(template);
        return new(offset, source.Length, source, template, refs);
    }

    static BlockMarker ParseBlockTag(string source, string templateName, string partUri)
    {
        if (!BlockTagParser.TryParse(source, out var tag, out var expression))
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Malformed block tag: {source}",
                partUri,
                source);
        }

        var expr = expression.IsEmpty ? null : expression.ToString();

        switch (tag)
        {
            case "for":
                return BuildForTag(source, expr, templateName, partUri);
            case "endfor":
                return new(BlockTagKind.EndFor, source, null, null, null);
            case "if":
                return BuildIfTag(source, expr, templateName, partUri);
            case "elsif":
            case "elseif":
                return BuildElsifTag(source, expr, templateName, partUri);
            case "else":
                return new(BlockTagKind.Else, source, null, null, null);
            case "endif":
                return new(BlockTagKind.EndIf, source, null, null, null);
            default:
                throw new ParchmentRegistrationException(
                    templateName,
                    $"Unsupported block tag '{tag}'. Supported: for, endfor, if, elsif, else, endif",
                    partUri,
                    source);
        }
    }

    static BlockMarker BuildForTag(string source, string? expression, string templateName, string partUri)
    {
        if (expression == null)
        {
            throw new ParchmentRegistrationException(
                templateName,
                "{% for %} tag is missing its loop expression",
                partUri,
                source);
        }

        // Parse `{% for x in y %}{% endfor %}` so Fluid yields a real ForStatement we can pull
        // Identifier and Source off — no manual regex of the loop variable, no manual reflection.
        var liquid = $"{{% for {expression} %}}{{% endfor %}}";
        if (!SharedFluid.Parser.TryParse(liquid, out var template, out var error))
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Failed to parse {{% for %}} tag: {error}",
                partUri,
                source);
        }

        var forStatement = ((Fluid.Parser.FluidTemplate)template).Statements
            .OfType<ForStatement>()
            .FirstOrDefault();
        if (forStatement == null)
        {
            throw new ParchmentRegistrationException(
                templateName,
                "{% for %} tag did not parse as a ForStatement",
                partUri,
                source);
        }

        return new(BlockTagKind.For, source, null, forStatement.Identifier, forStatement.Source);
    }

    static BlockMarker BuildIfTag(string source, string? expression, string templateName, string partUri) =>
        BuildConditional(BlockTagKind.If, source, expression, templateName, partUri);

    static BlockMarker BuildElsifTag(string source, string? expression, string templateName, string partUri) =>
        BuildConditional(BlockTagKind.ElsIf, source, expression, templateName, partUri);

    static BlockMarker BuildConditional(BlockTagKind kind, string source, string? expression, string templateName, string partUri)
    {
        if (expression == null)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"{{% {kind.ToString().ToLowerInvariant()} %}} tag is missing its condition",
                partUri,
                source);
        }

        // Parse `{% if cond %}{% endif %}` so Fluid yields a real IfStatement we can pull
        // Condition off — no string-comparison "true"/"false" trick at render time.
        var liquid = $"{{% if {expression} %}}{{% endif %}}";
        if (!SharedFluid.Parser.TryParse(liquid, out var template, out var error))
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Failed to parse {{% if %}} tag: {error}",
                partUri,
                source);
        }

        var ifStatement = ((Fluid.Parser.FluidTemplate)template).Statements
            .OfType<IfStatement>()
            .FirstOrDefault();
        if (ifStatement == null)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"{{% {kind.ToString().ToLowerInvariant()} %}} tag did not parse as an IfStatement",
                partUri,
                source);
        }

        return new(kind, source, ifStatement.Condition, null, null);
    }
}
