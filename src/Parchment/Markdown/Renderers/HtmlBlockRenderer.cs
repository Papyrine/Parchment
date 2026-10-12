class HtmlBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HtmlBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, HtmlBlock block)
    {
        // Comment-only blocks (snippet markers, TODOs, authoring notes) would otherwise round-trip
        // through the HTML converter as empty paragraphs and bleed visible whitespace into the docx.
        if (block.Type == HtmlBlockType.Comment)
        {
            return;
        }

        var html = block.Lines.ToString();
        var settings = renderer.ImagePolicies.BuildSettings(headingOffset: renderer.HeadingOffset);
        var elements = WordHtmlConverter.ToElements(html, renderer.MainPart, settings);
        renderer.MayHoldBookmarks = true;
        renderer.MayHoldSections = true;
        var indent = renderer.CurrentIndent;
        foreach (var element in elements)
        {
            if (indent > 0)
            {
                ApplyIndent(renderer, element, indent);
            }

            renderer.AddBlock(element);
        }
    }

    static void ApplyIndent(OpenXmlMarkdownRenderer renderer, OpenXmlElement element, int indent)
    {
        switch (element)
        {
            case Paragraph paragraph:
                ApplyParagraphIndent(renderer, paragraph, indent);
                break;
            case Table table:
                ApplyTableIndent(table, indent);
                break;
        }
    }

    // The converter hands back finished pPr and tblPr elements, which are strict schema sequences:
    // w:ind sits before w:jc and w:rPr, w:tblInd before w:tblBorders and w:tblLook. Appending put
    // the indent last, which is invalid as soon as the html carried an alignment, a border or a
    // list. The typed properties place the child where the schema wants it.
    static void ApplyParagraphIndent(OpenXmlMarkdownRenderer renderer, Paragraph paragraph, int indent)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        var existing = properties.Indentation;
        if (existing == null)
        {
            // A direct w:ind replaces the left indent a numbering level supplies, so an html list
            // starts from its level's indent; otherwise its markers land left of the container.
            properties.Indentation = new()
            {
                Left = (LevelIndent(renderer, properties) + indent).ToString(CultureInfo.InvariantCulture)
            };
            return;
        }

        existing.Left = (ParseIndent(existing.Left?.Value) + indent).ToString(CultureInfo.InvariantCulture);
    }

    static int LevelIndent(OpenXmlMarkdownRenderer renderer, ParagraphProperties properties)
    {
        var numberingProperties = properties.NumberingProperties;
        var numId = numberingProperties?.NumberingId?.Val?.Value;
        var numbering = renderer.MainPart.NumberingDefinitionsPart?.Numbering;
        if (numId == null ||
            numbering == null)
        {
            return 0;
        }

        var ilvl = numberingProperties!.NumberingLevelReference?.Val?.Value ?? 0;
        var abstractId = numbering
            .Elements<NumberingInstance>()
            .FirstOrDefault(_ => _.NumberID?.Value == numId)
            ?.AbstractNumId?.Val?.Value;
        if (abstractId == null)
        {
            return 0;
        }

        var level = numbering
            .Elements<AbstractNum>()
            .FirstOrDefault(_ => _.AbstractNumberId?.Value == abstractId)
            ?.Elements<Level>()
            .FirstOrDefault(_ => _.LevelIndex?.Value == ilvl);
        if (level == null)
        {
            return 0;
        }

        // A level's w:pPr is PreviousParagraphProperties once loaded from a package, but the
        // converter builds the levels it adds in memory with a plain ParagraphProperties.
        var indentation = level.PreviousParagraphProperties?.Indentation ??
                          level.GetFirstChild<ParagraphProperties>()?.Indentation;
        return ParseIndent(indentation?.Left?.Value);
    }

    static int ParseIndent(string? value)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return parsed;
        }

        return 0;
    }

    static void ApplyTableIndent(Table table, int indent)
    {
        var properties = table.GetFirstChild<TableProperties>();
        if (properties == null)
        {
            properties = new();
            table.InsertAt(properties, 0);
        }

        var existing = properties.TableIndentation;
        if (existing == null)
        {
            properties.TableIndentation = new()
            {
                Width = indent,
                Type = TableWidthUnitValues.Dxa
            };
            return;
        }

        existing.Width = (existing.Width?.Value ?? 0) + indent;
        existing.Type = TableWidthUnitValues.Dxa;
    }
}
