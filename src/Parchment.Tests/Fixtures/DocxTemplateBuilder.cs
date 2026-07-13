/// <summary>
/// Builds simple docx templates in-memory for test fixtures. The single <c>content</c> string is
/// split into paragraphs on blank lines — a line consisting only of whitespace separates two
/// paragraphs. Paragraph text may include liquid tokens.
/// </summary>
static class DocxTemplateBuilder
{
    public static MemoryStream Build(string content = "")
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            AddStyles(mainPart);

            foreach (var text in SplitParagraphs(content))
            {
                body.Append(BuildParagraph(text));
            }

            body.Append(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    },
                    new PageMargin
                    {
                        Top = 500,
                        Right = 500,
                        Bottom = 500,
                        Left = 500,
                        Header = 720,
                        Footer = 720
                    },
                    new PageBorders(
                        new TopBorder
                        {
                            Val = BorderValues.Single,
                            Size = 4,
                            Color = "000000",
                            Space = 0
                        },
                        new LeftBorder
                        {
                            Val = BorderValues.Single,
                            Size = 4,
                            Color = "000000",
                            Space = 0
                        },
                        new BottomBorder
                        {
                            Val = BorderValues.Single,
                            Size = 4,
                            Color = "000000",
                            Space = 0
                        },
                        new RightBorder
                        {
                            Val = BorderValues.Single,
                            Size = 4,
                            Color = "000000",
                            Space = 0
                        })
                    {
                        OffsetFrom = PageBorderOffsetValues.Page,
                        Display = PageBorderDisplayValues.AllPages,
                        ZOrder = PageBorderZOrderValues.Front
                    }));
        }

        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Builds a template whose body is a single one-row, two-column table: a label cell and a
    /// value cell containing <paramref name="valueCellText"/> as its only paragraph (the
    /// whole-cell editable-range shape). <paramref name="extraValueCellParagraph"/> appends a
    /// second paragraph to the value cell to exercise the shared-cell fallback.
    /// </summary>
    public static MemoryStream BuildWithTable(
        string labelText,
        string valueCellText,
        bool extraValueCellParagraph = false)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            AddStyles(mainPart);

            var valueCell = new TableCell(BuildParagraph(valueCellText));
            if (extraValueCellParagraph)
            {
                valueCell.AppendChild(BuildParagraph("sibling"));
            }

            body.Append(
                new Table(
                    new TableProperties(
                        new TableWidth
                        {
                            Type = TableWidthUnitValues.Auto
                        }),
                    new TableGrid(
                        new GridColumn(),
                        new GridColumn()),
                    new TableRow(
                        new TableCell(BuildParagraph(labelText)),
                        valueCell)));
            body.Append(new SectionProperties(new PageSize
            {
                Width = 6500,
                Height = 8000
            }));
        }

        stream.Position = 0;
        return stream;
    }

    static IEnumerable<string> SplitParagraphs(string content)
    {
        if (string.IsNullOrEmpty(content))
        {
            yield break;
        }

        var normalized = content.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var current = new StringBuilder();
        foreach (var line in lines)
        {
            if (line.TrimStart().StartsWith("//"))
            {
                continue;
            }

            if (line.Length == 0 || string.IsNullOrWhiteSpace(line))
            {
                if (current.Length > 0)
                {
                    yield return current.ToString();
                    current.Clear();
                }
                continue;
            }

            if (current.Length > 0)
            {
                current.Append('\n');
            }
            current.Append(line);
        }

        if (current.Length > 0)
        {
            yield return current.ToString();
        }
    }

    static Paragraph BuildParagraph(string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    static void AddStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        styles.Append(BuildStyle("Normal", StyleValues.Paragraph, isDefault: true));
        for (var i = 1; i <= 6; i++)
        {
            styles.Append(BuildStyle($"Heading{i}", StyleValues.Paragraph));
        }

        styles.Append(BuildStyle("ListParagraph", StyleValues.Paragraph));
        styles.Append(BuildStyle("Quote", StyleValues.Paragraph));
        styles.Append(BuildStyle("Code", StyleValues.Paragraph));
        styles.Append(BuildStyle("Hyperlink", StyleValues.Character));

        stylesPart.Styles = styles;
    }

    static Style BuildStyle(string id, StyleValues type, bool isDefault = false)
    {
        var style = new Style
        {
            Type = type,
            StyleId = id
        };
        style.Append(new StyleName { Val = id });
        if (isDefault)
        {
            style.Default = true;
        }

        return style;
    }
}
