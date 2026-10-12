/// <summary>
/// Fills a table-of-contents field with the entries it describes, so the document reaches its final
/// length before anything measures it.
/// </summary>
/// <remarks>
/// Word builds these entries itself when it updates the field. Building them here instead is what
/// makes a resolved table of contents possible at all: entries added after pagination would push
/// every heading below them onto a later page and invalidate the numbers just measured, whereas a
/// number written into a line that already exists reflows nothing.
/// <para>
/// The entries stay inside the field rather than beside it. A reader who refreshes the field then
/// gets them replaced, which is the point — outside it, refreshing would insert a second table.
/// </para>
/// </remarks>
static partial class WordTableOfContentsBuilder
{
    // Stands in for the page number until the resolver reports one. A single digit keeps the entry
    // the width it will end up, so filling the real number in cannot rewrap the line.
    internal const string PagePlaceholder = "1";

    public static void WriteEntries(Body body, WordFieldRegion field, BookmarkIds ids)
    {
        var maxLevel = ParseMaxLevel(field.Instruction);
        var headings = CollectHeadings(body, maxLevel);
        if (headings.Count == 0)
        {
            // Nothing to list. The placeholder stays, and so does the request for Word to try again.
            return;
        }

        var bookmarks = new WordBookmarkAllocator(ids);
        var tabStop = TextWidth(SectionOf(body, field.Paragraph));
        var entries = new List<Paragraph>();
        foreach (var (paragraph, level, text) in headings)
        {
            entries.Add(BuildEntry(bookmarks.Ensure(paragraph), level, text, tabStop));
        }

        // The field's markers move onto the entries: begin/instruction/separate lead the first, and
        // end closes the last, which is how Word writes a table of contents spanning paragraphs.
        var source = field.Paragraph;
        var leading = source.Elements<Run>()
            .TakeWhile(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Separate)
            .Concat(source.Elements<Run>()
                .Where(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate)
                .Take(1))
            .ToList();
        var end = source.Elements<Run>()
            .LastOrDefault(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End);

        // After the entry's properties rather than at the front of it: w:pPr has to stay the
        // paragraph's first child, and Word discards the whole of it when a run comes first — so the
        // first entry alone would lose its TOC style and the tab stop its dot leader hangs off.
        var anchor = entries[0].ParagraphProperties!;
        for (var i = leading.Count - 1; i >= 0; i--)
        {
            leading[i].Remove();
            anchor.InsertAfterSelf(leading[i]);
        }

        if (end != null)
        {
            end.Remove();
            entries[^1].Append(end);
        }

        foreach (var entry in entries)
        {
            source.InsertBeforeSelf(entry);
        }

        source.Remove();
    }

    static Paragraph BuildEntry(string bookmark, int level, string text, int tabStop)
    {
        var properties = new ParagraphProperties
        {
            ParagraphStyleId = new()
            {
                Val = $"TOC{level}"
            }
        };
        properties.Append(
            new Tabs(
                new TabStop
                {
                    Val = TabStopValues.Right,
                    Leader = TabStopLeaderCharValues.Dot,
                    Position = tabStop
                }));

        var hyperlink = new Hyperlink
        {
            Anchor = bookmark
        };
        hyperlink.Append(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
        hyperlink.Append(new Run(new TabChar()));
        hyperlink.Append(new Run(new Text(PagePlaceholder)));

        return new(properties, hyperlink);
    }

    static List<(Paragraph Paragraph, int Level, string Text)> CollectHeadings(Body body, int maxLevel)
    {
        var found = new List<(Paragraph, int, string)>();
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (style == null ||
                !style.StartsWith("Heading", StringComparison.Ordinal) ||
                !int.TryParse(style.AsSpan("Heading".Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out var level) ||
                level > maxLevel)
            {
                continue;
            }

            var text = string.Concat(paragraph.Descendants<Text>().Select(_ => _.Text)).Trim();
            if (text.Length == 0)
            {
                continue;
            }

            found.Add((paragraph, level, text));
        }

        return found;
    }

    static int ParseMaxLevel(string instruction)
    {
        // \o "1-3"
        var match = ParseMaxLevelRegex()
            .Match(instruction);
        if (match.Success &&
            int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var level))
        {
            return level;
        }

        return 3;
    }

    // The section the field sits in, which is not the document's last once [SECTION] has split it: a
    // w:sectPr describes the section that ends at its paragraph, so the one in force here is the
    // first to close at or below the field, and the document-final one only when none does. Looked
    // for among the body's own children, where a section break has to be, starting from whichever
    // of them holds the field — the paragraph itself, or the table it is in a cell of. Read after
    // SectionBreaks.Apply, which is what turns each marker into the full section measured here.
    static SectionProperties? SectionOf(Body body, Paragraph paragraph)
    {
        OpenXmlElement block = paragraph;
        while (block.Parent != null &&
               !ReferenceEquals(block.Parent, body))
        {
            block = block.Parent;
        }

        for (var current = block; current != null; current = current.NextSibling())
        {
            if (current is Paragraph { ParagraphProperties.SectionProperties: { } section })
            {
                return section;
            }
        }

        return body.Elements<SectionProperties>().LastOrDefault();
    }

    // Where the page number sits: hard against the right margin, so the dot leader spans the gap.
    static int TextWidth(SectionProperties? section)
    {
        var size = section?.GetFirstChild<PageSize>();
        var margin = section?.GetFirstChild<PageMargin>();
        if (size?.Width?.Value is not { } width)
        {
            // A4 less the margins Word defaults to.
            return 9026;
        }

        var left = margin?.Left?.Value ?? 1440;
        var right = margin?.Right?.Value ?? 1440;
        return (int)width - (int)left - (int)right;
    }

    [GeneratedRegex(
        """
        \\o\s+"\d+-(\d+)"
        """)]
    private static partial Regex ParseMaxLevelRegex();
}
