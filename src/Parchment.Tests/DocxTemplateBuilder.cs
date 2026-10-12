using Kit = OpenXmlKit.Word;

/// <summary>
/// Builds simple docx templates in-memory for test fixtures. The single <c>content</c> string is
/// split into paragraphs on blank lines — a line consisting only of whitespace separates two
/// paragraphs. Paragraph text may include liquid tokens.
/// </summary>
/// <remarks>
/// OpenXmlKit supplies the package and the styles; the body content and the section stay raw on
/// purpose, so what this produces is byte-for-byte what it always did. The build API would have
/// been an improvement to read and a change to the fixtures — a newline inside a <c>w:t</c> would
/// become a <c>w:br</c>, a table would gain the trailing paragraph Word wants after one — and a
/// fixture is the wrong place to accept either.
/// </remarks>
static class DocxTemplateBuilder
{
    public static MemoryStream Build(string content = "")
    {
        var stream = new MemoryStream();
        using (var document = Kit.Document.Create(stream))
        {
            AddStyles(document.Styles);
            var body = document.Body.ToOpenXml();

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
    /// second paragraph to the value cell to exercise the shared-cell fallback;
    /// <paramref name="bodyParagraphText"/> prepends a body paragraph before the table.
    /// </summary>
    public static MemoryStream BuildWithTable(
        string labelText,
        string valueCellText,
        bool extraValueCellParagraph = false,
        string? bodyParagraphText = null)
    {
        var stream = new MemoryStream();
        using (var document = Kit.Document.Create(stream))
        {
            AddStyles(document.Styles);
            var body = document.Body.ToOpenXml();

            if (bodyParagraphText != null)
            {
                body.Append(BuildParagraph(bodyParagraphText));
            }

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

    // Stubs rather than real definitions: these exist so a template names a style that resolves,
    // not so it renders like anything.
    static void AddStyles(Kit.Styles styles)
    {
        styles.Add(Kit.StyleKind.Paragraph, "Normal").IsDefault = true;
        for (var i = 1; i <= 6; i++)
        {
            styles.Add(Kit.StyleKind.Paragraph, $"Heading{i}");
        }

        styles.Add(Kit.StyleKind.Paragraph, "ListParagraph");
        styles.Add(Kit.StyleKind.Paragraph, "Quote");
        styles.Add(Kit.StyleKind.Paragraph, "Code");
        styles.Add(Kit.StyleKind.Character, "Hyperlink");
    }
}
