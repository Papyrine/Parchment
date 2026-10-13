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
/// <para>
/// The converter mimics an element's default look with direct run properties: every run of a
/// heading carries <c>w:b</c>, every run of a link carries <c>w:u</c> (and the link colour). That
/// formatting is implied by the <c>h1</c>–<c>h6</c> / <c>a</c> being emitted, so it is not
/// reported again as <c>strong</c> / <c>u</c> — otherwise unedited content would not extract to
/// the html it was rendered from. Only a departure from the default is reported, and in the form
/// the converter reads back: <c>font-weight:normal</c> for a heading run whose bold is explicitly
/// off, <c>text-decoration:none</c> on a link that is not underlined throughout.
/// </para>
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
        var heading = HeadingTag(paragraph);
        var tag = heading ?? "p";
        builder.Append('<').Append(tag).Append('>');
        EmitInline(builder, paragraph, mainPart, heading != null);
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
            EmitInline(builder, paragraph, mainPart, false);
            currentLevel = level;
            index++;
        }

        while (openTags.Count > 0)
        {
            builder.Append("</li></").Append(openTags.Pop()).Append('>');
        }

        return index;
    }

    // linkUnderlined is null outside a link. Inside one it says whether the link is underlined
    // throughout, which is what decides how a run of it reports its own underline.
    static void EmitInline(StringBuilder builder, OpenXmlElement container, MainDocumentPart mainPart, bool inHeading, bool? linkUnderlined = null)
    {
        foreach (var child in EditableContent.Inline(container))
        {
            switch (child)
            {
                case Run run:
                    EmitRun(builder, run, inHeading, ReportsUnderline(run, linkUnderlined));
                    break;
                case Hyperlink hyperlink:
                    EmitHyperlink(builder, hyperlink, mainPart, inHeading);
                    break;
                default:
                    // Math has no HTML form in this subset: keep what it says.
                    builder.Append(Escape(child.InnerText));
                    break;
            }
        }
    }

    static void EmitHyperlink(StringBuilder builder, Hyperlink hyperlink, MainDocumentPart mainPart, bool inHeading)
    {
        var href = ResolveHref(hyperlink, mainPart);
        if (href == null)
        {
            // Not emitted as a link, so nothing implies the underline.
            EmitInline(builder, hyperlink, mainPart, inHeading);
            return;
        }

        // An <a> is underlined by default. Underlined throughout, the link says it all; otherwise
        // the default is switched off on the <a> and the runs that do keep an underline say so.
        var underlined = EditableContent.Inline(hyperlink)
            .OfType<Run>()
            .All(_ => !HasContent(_) || IsLinkUnderlined(_));
        builder.Append("<a href=\"").Append(EscapeAttribute(href)).Append('"');
        if (!underlined)
        {
            builder.Append(" style=\"text-decoration:none\"");
        }

        builder.Append('>');
        EmitInline(builder, hyperlink, mainPart, inHeading, underlined);
        builder.Append("</a>");
    }

    static bool ReportsUnderline(Run run, bool? linkUnderlined)
    {
        if (linkUnderlined == null)
        {
            return HasUnderline(run.RunProperties);
        }

        return !linkUnderlined.Value &&
               IsLinkUnderlined(run);
    }

    static bool HasContent(Run run) =>
        run.Elements().Any(_ => RunCharacters.For(_) != null);

    // A link the converter wrote carries a direct w:u; one inserted in Word gets its underline
    // from the Hyperlink character style instead. Removing the underline in Word drops the direct
    // w:u, or writes w:u w:val="none" over the style.
    static bool IsLinkUnderlined(Run run)
    {
        var properties = run.RunProperties;
        if (properties?.Underline != null)
        {
            return HasUnderline(properties);
        }

        var style = properties?.RunStyle?.Val?.Value;
        return string.Equals(style, "Hyperlink", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(style, "FollowedHyperlink", StringComparison.OrdinalIgnoreCase);
    }

    static void EmitRun(StringBuilder builder, Run run, bool inHeading, bool underline)
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

        // A heading is bold by default, so bold says nothing there. A run without w:b is left as
        // the heading's own look too (a paragraph given a heading style in Word has none); only an
        // explicit off is a visible departure.
        var boldOff = inHeading && IsOff(properties?.Bold);
        if (boldOff)
        {
            builder.Append("<span style=\"font-weight:normal\">");
        }

        if (!inHeading &&
            IsOn(properties?.Bold))
        {
            tags.Add("strong");
        }

        if (IsOn(properties?.Italic))
        {
            tags.Add("em");
        }

        if (underline)
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

        if (boldOff)
        {
            builder.Append("</span>");
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

    static bool IsOff(OnOffType? toggle) =>
        toggle?.Val != null &&
        !toggle.Val.Value;

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
