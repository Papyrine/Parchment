/// <summary>
/// Turns a paragraph whose whole content is <c>[SECTION]</c> into a Word section break, so a markdown
/// template can change page orientation — and margins with it — partway down a document.
/// </summary>
/// <remarks>
/// The spelling follows <see cref="TableOfContents"/>: with no <c>(url)</c> or <c>[label]</c> after it,
/// <c>[SECTION]</c> parses as the literal text "[SECTION]", so the marker costs no syntax and a
/// template that renders to html elsewhere degrades to showing it rather than breaking. A marker
/// whose <c>orientation</c> is missing or unrecognised is not treated as a marker at all and renders
/// as that literal text, which is what makes a typo visible in the output instead of a section break
/// that quietly does nothing.
/// <para>
/// Writing happens in two halves because the two halves know different things. <see cref="TryWrite"/>
/// runs inside the markdown walk and can only record what the template asked for, so it emits a
/// paragraph carrying a <c>w:sectPr</c> holding just that. <see cref="Apply"/> runs in
/// <c>RegisteredMarkdownTemplate.Render</c>, where the style source's own <c>w:sectPr</c> is in hand,
/// and rewrites each of those into a full section derived from it — so headers, footers, page borders
/// and everything else the template set carry into every section rather than being dropped for the
/// two or three properties the marker named.
/// </para>
/// <para>
/// The halves also disagree about direction. <c>[SECTION]</c> reads as "from here on", which is how
/// the author thinks about it, but a <c>w:sectPr</c> inside a paragraph's <c>w:pPr</c> describes the
/// section that <em>ends</em> at that paragraph. So <see cref="Apply"/> walks the markers writing each
/// one the settings that were in force <em>before</em> it, and the settings the last marker declared
/// land on the document-final <c>w:sectPr</c> instead.
/// </para>
/// </remarks>
static class SectionBreaks
{
    const string Marker = "[SECTION]";

    // A4, matching Word's own default, for a style source whose section properties do not say.
    const uint DefaultWidth = 11906;
    const uint DefaultHeight = 16838;

    // The children of w:sectPr that follow w:pgSz, in schema order. Only needed to place a w:pgSz
    // into a style source that carries none — Word reads sectPr positionally, so appending would
    // put it after elements it has to precede.
    static readonly Type[] afterPageSize =
    [
        typeof(PageMargin),
        typeof(PaperSource),
        typeof(PageBorders),
        typeof(LineNumberType),
        typeof(PageNumberType),
        typeof(Columns),
        typeof(FormProtection),
        typeof(VerticalTextAlignmentOnPage),
        typeof(NoEndnote),
        typeof(TitlePage),
        typeof(TextDirection),
        typeof(BiDi),
        typeof(GutterOnRight),
        typeof(DocGrid),
        typeof(PrinterSettingsReference)
    ];

    public static bool TryWrite(OpenXmlMarkdownRenderer renderer, ParagraphBlock block, ParagraphProperties? properties)
    {
        if (!IsMarker(block))
        {
            return false;
        }

        var attributes = block.TryGetAttributes()?.Properties;
        if (ResolveOrientation(attributes) is not { } orientation)
        {
            return false;
        }

        var declared = new SectionProperties(
            new PageSize
            {
                Orient = orientation
            });

        if (ResolveMargins(attributes) is { } margins)
        {
            declared.AppendChild(margins);
        }

        properties ??= new();
        properties.SectionProperties = declared;
        renderer.AddBlock(new Paragraph(properties));
        return true;
    }

    /// <summary>
    /// Rewrite every marker written by <see cref="TryWrite"/> into a real section derived from
    /// <paramref name="baseSection"/>, which is mutated in place to become the document-final section.
    /// </summary>
    public static void Apply(Body body, SectionProperties? baseSection, string name)
    {
        var markers = body
            .Descendants<Paragraph>()
            .Where(_ => _.ParagraphProperties?.SectionProperties != null)
            .ToList();

        if (markers.Count == 0)
        {
            return;
        }

        if (baseSection == null)
        {
            throw new ParchmentRenderException(
                name,
                $"'{Marker}' needs a style source to derive its page setup from, and this template has none.");
        }

        // The settings in force at the top of the document. Each marker replaces this with what it
        // declared, so the marker after it — or the document-final section — inherits from there
        // rather than from the style source.
        var current = (SectionProperties) baseSection.CloneNode(true);

        for (var index = 0; index < markers.Count; index++)
        {
            var marker = markers[index];
            if (!ReferenceEquals(marker.Parent, body))
            {
                throw new ParchmentRenderException(
                    name,
                    $"'{Marker}' has to be a paragraph of its own at the top level of the document. " +
                    "One inside a list, a table cell or a block quote is not a section break Word can express.");
            }

            var properties = marker.ParagraphProperties!;
            var declared = properties.SectionProperties!;
            var next = Derive(current, declared);

            // Page numbering restarts wherever w:pgNumType carries a start, so only the first
            // section keeps the style source's. Every later one continues from the section before
            // it, which is what inserting a section break in Word gives you.
            if (index > 0)
            {
                StripPageNumberStart(current);
            }

            properties.SectionProperties = current;
            current = next;
        }

        StripPageNumberStart(current);

        // The caller appends baseSection after this returns, so the last declared settings have to
        // land on that instance rather than on a replacement for it.
        baseSection.RemoveAllChildren();
        foreach (var child in current.ChildElements.ToList())
        {
            baseSection.AppendChild(child.CloneNode(true));
        }
    }

    // An unresolved "[SECTION]" reaches the renderer as plain text, but not necessarily as one run of
    // it — the link parser splits the brackets off from the word between them. So the test is that
    // the paragraph holds text and nothing else, and that the text is the marker.
    static bool IsMarker(ParagraphBlock block)
    {
        var child = block.Inline?.FirstChild;
        if (child == null)
        {
            return false;
        }

        var builder = new StringBuilder();
        while (child != null)
        {
            if (child is not LiteralInline literal)
            {
                return false;
            }

            builder.Append(literal.Content.AsSpan());
            child = child.NextSibling;
        }

        return builder.ToString().Trim() == Marker;
    }

    static PageOrientationValues? ResolveOrientation(IEnumerable<KeyValuePair<string, string?>>? attributes)
    {
        if (attributes == null)
        {
            return null;
        }

        foreach (var (key, value) in attributes)
        {
            if (!string.Equals(key, "orientation", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(value, "landscape", StringComparison.OrdinalIgnoreCase))
            {
                return PageOrientationValues.Landscape;
            }

            if (string.Equals(value, "portrait", StringComparison.OrdinalIgnoreCase))
            {
                return PageOrientationValues.Portrait;
            }

            return null;
        }

        return null;
    }

    // margins=top,right,bottom,left in twips, in the order css writes them. Unparseable margins are
    // ignored rather than rejected, the way an unparseable [TOC] level is: the section break itself
    // is still what the author asked for, and the page keeps the margins it already had.
    static PageMargin? ResolveMargins(IEnumerable<KeyValuePair<string, string?>>? attributes)
    {
        if (attributes == null)
        {
            return null;
        }

        foreach (var (key, value) in attributes)
        {
            if (!string.Equals(key, "margins", StringComparison.OrdinalIgnoreCase) ||
                value == null)
            {
                continue;
            }

            var parts = value.Split(',');
            if (parts.Length != 4)
            {
                return null;
            }

            var twips = new int[4];
            for (var index = 0; index < 4; index++)
            {
                if (!int.TryParse(parts[index].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out twips[index]))
                {
                    return null;
                }
            }

            return new()
            {
                Top = twips[0],
                Right = (uint) twips[1],
                Bottom = twips[2],
                Left = (uint) twips[3]
            };
        }

        return null;
    }

    static SectionProperties Derive(SectionProperties current, SectionProperties declared)
    {
        var next = (SectionProperties) current.CloneNode(true);
        ApplyOrientation(next, declared.GetFirstChild<PageSize>()!.Orient!.Value);

        if (declared.GetFirstChild<PageMargin>() is { } margins)
        {
            var target = next.GetFirstChild<PageMargin>();
            if (target == null)
            {
                next.AppendChild((PageMargin) margins.CloneNode(true));
            }
            else
            {
                // Only the four page edges: header, footer and gutter distances are the style
                // source's to set, and a template changing orientation has no opinion on them.
                target.Top = margins.Top;
                target.Right = margins.Right;
                target.Bottom = margins.Bottom;
                target.Left = margins.Left;
            }
        }

        return next;
    }

    // Read as sides rather than as width and height, so the result is the requested orientation
    // whichever one the style source was in.
    static void ApplyOrientation(SectionProperties sectPr, PageOrientationValues orientation)
    {
        var size = EnsurePageSize(sectPr);
        var width = size.Width?.Value ?? DefaultWidth;
        var height = size.Height?.Value ?? DefaultHeight;
        var shortSide = Math.Min(width, height);
        var longSide = Math.Max(width, height);

        if (orientation == PageOrientationValues.Landscape)
        {
            size.Width = longSide;
            size.Height = shortSide;
            size.Orient = PageOrientationValues.Landscape;
            return;
        }

        size.Width = shortSide;
        size.Height = longSide;
        // Portrait is what Word assumes, so saying so adds an attribute that changes nothing.
        size.Orient = null;
    }

    static PageSize EnsurePageSize(SectionProperties sectPr)
    {
        if (sectPr.GetFirstChild<PageSize>() is { } existing)
        {
            return existing;
        }

        var size = new PageSize
        {
            Width = DefaultWidth,
            Height = DefaultHeight
        };

        foreach (var child in sectPr.ChildElements)
        {
            if (afterPageSize.Contains(child.GetType()))
            {
                sectPr.InsertBefore(size, child);
                return size;
            }
        }

        sectPr.AppendChild(size);
        return size;
    }

    // The element carries a format as well, so it is emptied of its start rather than removed.
    static void StripPageNumberStart(SectionProperties sectPr)
    {
        if (sectPr.GetFirstChild<PageNumberType>() is not { } pageNumbers)
        {
            return;
        }

        pageNumbers.Start = null;
        if (!pageNumbers.HasAttributes)
        {
            pageNumbers.Remove();
        }
    }
}
