namespace Parchment;

/// <summary>
/// Serializes the block content of an editable-HTML control (<see cref="EditableFieldKind.Html"/>)
/// back into an HTML string — the reverse of OpenXmlHtml's <c>WordHtmlConverter.ToElements</c>.
/// It reads <em>standard</em> WordprocessingML (paragraph styles, <c>w:numPr</c> lists, run
/// <c>rPr</c> emphasis, hyperlinks), so it inverts both the original render and any edits Word
/// itself wrote. Supported subset: paragraphs, headings, bold/italic/underline/strike,
/// superscript/subscript, hyperlinks, line breaks, tabs, and bullet/ordered lists (with nesting).
/// Anything outside the subset degrades to its text content; <see cref="EditableContent"/> decides
/// which paragraphs and runs the control holds.
/// </summary>
static class OpenXmlHtmlSerializer
{
    public static string ToHtml(OpenXmlElement content, MainDocumentPart mainPart)
    {
        var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
        // Every paragraph the control holds, however deeply Word wrapped it. A table has no HTML
        // round-trip in this subset, so it arrives here as its cell paragraphs and no text is
        // dropped.
        var blocks = EditableContent.Paragraphs(content).ToList<OpenXmlElement>();
        var builder = new StringBuilder();

        for (var i = 0; i < blocks.Count;)
        {
            var paragraph = (Paragraph) blocks[i];
            if (IsListItem(paragraph, out _, out _))
            {
                i = EmitList(builder, blocks, i, numbering, mainPart);
                continue;
            }

            EmitParagraph(builder, paragraph, mainPart);
            i++;
        }

        return builder.ToString();
    }

    static void EmitParagraph(StringBuilder builder, Paragraph paragraph, MainDocumentPart mainPart)
    {
        var tag = HeadingTag(paragraph) ?? "p";
        builder.Append('<').Append(tag).Append('>');
        EmitInline(builder, paragraph, mainPart);
        builder.Append("</").Append(tag).Append('>');
    }

    static int EmitList(StringBuilder builder, List<OpenXmlElement> blocks, int start, Numbering? numbering, MainDocumentPart mainPart)
    {
        var openTags = new Stack<string>();
        var currentLevel = -1;
        var index = start;

        while (index < blocks.Count &&
               blocks[index] is Paragraph paragraph &&
               IsListItem(paragraph, out var numberingId, out var level))
        {
            if (level > currentLevel)
            {
                // Open one list per new depth; a deeper list nests inside the still-open <li>.
                for (var open = currentLevel + 1; open <= level; open++)
                {
                    var tag = ListTag(numbering, numberingId, open);
                    builder.Append('<').Append(tag).Append('>');
                    openTags.Push(tag);
                }
            }
            else
            {
                while (currentLevel > level)
                {
                    builder.Append("</li></").Append(openTags.Pop()).Append('>');
                    currentLevel--;
                }

                builder.Append("</li>");

                // A sibling of the other kind (bullet next to ordered) is a new list, not another
                // item. Only the kind splits: Word hands out a fresh w:numId for bullets a user
                // sees as one list, so the numbering instance alone is not a boundary.
                var tag = ListTag(numbering, numberingId, level);
                if (tag != openTags.Peek())
                {
                    builder.Append("</").Append(openTags.Pop()).Append('>');
                    builder.Append('<').Append(tag).Append('>');
                    openTags.Push(tag);
                }
            }

            builder.Append("<li>");
            EmitInline(builder, paragraph, mainPart);
            currentLevel = level;
            index++;
        }

        while (openTags.Count > 0)
        {
            builder.Append("</li></").Append(openTags.Pop()).Append('>');
        }

        return index;
    }

    static void EmitInline(StringBuilder builder, OpenXmlElement container, MainDocumentPart mainPart)
    {
        foreach (var child in EditableContent.Inline(container))
        {
            switch (child)
            {
                case Run run:
                    EmitRun(builder, run);
                    break;
                case Hyperlink hyperlink:
                    EmitHyperlink(builder, hyperlink, mainPart);
                    break;
                default:
                    // Math has no HTML form in this subset: keep what it says.
                    builder.Append(Escape(child.InnerText));
                    break;
            }
        }
    }

    static void EmitHyperlink(StringBuilder builder, Hyperlink hyperlink, MainDocumentPart mainPart)
    {
        var href = ResolveHref(hyperlink, mainPart);
        if (href != null)
        {
            builder.Append("<a href=\"").Append(EscapeAttribute(href)).Append("\">");
        }

        EmitInline(builder, hyperlink, mainPart);

        if (href != null)
        {
            builder.Append("</a>");
        }
    }

    static void EmitRun(StringBuilder builder, Run run)
    {
        var inner = new StringBuilder();
        foreach (var child in run.Elements())
        {
            AppendCharacters(inner, RunCharacters.For(child));
        }

        if (inner.Length == 0)
        {
            return;
        }

        var properties = run.RunProperties;
        var tags = new List<string>();
        if (IsOn(properties?.Bold))
        {
            tags.Add("strong");
        }

        if (IsOn(properties?.Italic))
        {
            tags.Add("em");
        }

        if (HasUnderline(properties))
        {
            tags.Add("u");
        }

        if (IsOn(properties?.Strike))
        {
            tags.Add("s");
        }

        // The OpenXML v3 "enum" value types are structs, not compile-time constants, so these are
        // == comparisons rather than a switch.
        var vertical = properties?.VerticalTextAlignment?.Val?.Value;
        if (vertical == VerticalPositionValues.Superscript)
        {
            tags.Add("sup");
        }
        else if (vertical == VerticalPositionValues.Subscript)
        {
            tags.Add("sub");
        }

        foreach (var tag in tags)
        {
            builder.Append('<').Append(tag).Append('>');
        }

        builder.Append(inner);

        for (var i = tags.Count - 1; i >= 0; i--)
        {
            builder.Append("</").Append(tags[i]).Append('>');
        }
    }

    // What OpenXmlHtml turns back into a w:tab. Written as the bare character a tab would fold to
    // a space when the html is rendered again, as html's own whitespace rules require.
    const string tabHtml = "<span style=\"white-space:pre\">&#9;</span>";

    // A line break and a tab are markup in html as they are elements in Word; everything else is
    // text.
    static void AppendCharacters(StringBuilder builder, string? characters)
    {
        if (characters == null)
        {
            return;
        }

        var start = 0;
        for (var index = 0; index < characters.Length; index++)
        {
            var character = characters[index];
            if (character is not ('\n' or '\t'))
            {
                continue;
            }

            builder.Append(Escape(characters[start..index]));
            if (character == '\n')
            {
                builder.Append("<br>");
            }
            else
            {
                builder.Append(tabHtml);
            }

            start = index + 1;
        }

        builder.Append(Escape(characters[start..]));
    }

    static bool IsListItem(Paragraph paragraph, out int numberingId, out int level)
    {
        numberingId = 0;
        level = 0;
        var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
        var id = numberingProperties?.NumberingId?.Val;
        if (id is not { HasValue: true } ||
            id.Value == 0)
        {
            return false;
        }

        numberingId = id.Value;
        level = numberingProperties!.NumberingLevelReference?.Val?.Value ?? 0;
        return true;
    }

    static string ListTag(Numbering? numbering, int numberingId, int level)
    {
        if (numbering == null)
        {
            return "ul";
        }

        var abstractId = numbering
            .Elements<NumberingInstance>()
            .FirstOrDefault(_ => _.NumberID?.Value == numberingId)
            ?.AbstractNumId?.Val?.Value;
        if (abstractId == null)
        {
            return "ul";
        }

        var abstractNum = numbering
            .Elements<AbstractNum>()
            .FirstOrDefault(_ => _.AbstractNumberId?.Value == abstractId);
        var format = (abstractNum?
                .Elements<Level>()
                .FirstOrDefault(_ => _.LevelIndex?.Value == level) ??
            abstractNum?.Elements<Level>().FirstOrDefault())
            ?.NumberingFormat?.Val?.Value;

        if (format == NumberFormatValues.Bullet)
        {
            return "ul";
        }

        return "ol";
    }

    static string? HeadingTag(Paragraph paragraph)
    {
        var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (style == null ||
            !style.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) ||
            !int.TryParse(style.AsSpan("Heading".Length), out var level) ||
            level < 1)
        {
            return null;
        }

        return "h" + Math.Min(level, 6);
    }

    static string? ResolveHref(Hyperlink hyperlink, MainDocumentPart mainPart)
    {
        var anchor = hyperlink.Anchor?.Value;
        var target = ResolveTarget(hyperlink, mainPart);
        if (target == null)
        {
            // No relationship: a link to a bookmark in this document.
            if (string.IsNullOrEmpty(anchor))
            {
                return null;
            }

            return "#" + anchor;
        }

        // Word never keeps a fragment in the relationship target: on save it moves it to w:anchor
        // (or w:docLocation for a location that is not a bookmark), so the url is the two rejoined.
        // w:anchor wins when both are present, as it does in Word.
        var fragment = anchor;
        if (string.IsNullOrEmpty(fragment))
        {
            fragment = hyperlink.DocLocation?.Value;
        }

        if (string.IsNullOrEmpty(fragment) ||
            target.Contains('#'))
        {
            return target;
        }

        return target + "#" + fragment;
    }

    static string? ResolveTarget(Hyperlink hyperlink, MainDocumentPart mainPart)
    {
        var id = hyperlink.Id?.Value;
        if (id == null)
        {
            return null;
        }

        // OriginalString is the target as stored. Uri.ToString() would decode percent-escapes
        // (%20 to a space, %22 to a quote), handing back a different url than the one linked.
        return mainPart.HyperlinkRelationships.FirstOrDefault(_ => _.Id == id)?.Uri.OriginalString;
    }

    static bool IsOn(OnOffType? toggle) =>
        toggle != null &&
        (toggle.Val == null || toggle.Val.Value);

    static bool HasUnderline(RunProperties? properties)
    {
        var underline = properties?.Underline;
        return underline != null &&
               underline.Val?.Value != UnderlineValues.None;
    }

    static string Escape(string value) =>
        value
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;");

    static string EscapeAttribute(string value) =>
        Escape(value).Replace("\"", "&quot;");
}
