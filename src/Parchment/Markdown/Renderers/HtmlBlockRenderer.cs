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
                ApplyIndent(element, indent);
            }

            renderer.AddBlock(element);
        }
    }

    static void ApplyIndent(OpenXmlElement element, int indent)
    {
        switch (element)
        {
            case Paragraph paragraph:
                ApplyParagraphIndent(paragraph, indent);
                break;
            case Table table:
                ApplyTableIndent(table, indent);
                break;
        }
    }

    static void ApplyParagraphIndent(Paragraph paragraph, int indent)
    {
        paragraph.ParagraphProperties ??= new();
        var existing = paragraph.ParagraphProperties.GetFirstChild<Indentation>();
        if (existing == null)
        {
            paragraph.ParagraphProperties.Append(
                new Indentation
                {
                    Left = indent.ToString(CultureInfo.InvariantCulture)
                });
            return;
        }

        var current = int.TryParse(existing.Left?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0;
        existing.Left = (current + indent).ToString(CultureInfo.InvariantCulture);
    }

    static void ApplyTableIndent(Table table, int indent)
    {
        var properties = table.GetFirstChild<TableProperties>();
        if (properties == null)
        {
            properties = new();
            table.InsertAt(properties, 0);
        }

        var existing = properties.GetFirstChild<TableIndentation>();
        if (existing == null)
        {
            properties.Append(
                new TableIndentation
                {
                    Width = indent,
                    Type = TableWidthUnitValues.Dxa
                });
            return;
        }

        existing.Width = (existing.Width?.Value ?? 0) + indent;
        existing.Type = TableWidthUnitValues.Dxa;
    }
}
