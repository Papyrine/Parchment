/// <summary>
/// Turns a paragraph whose whole content is <c>[SECTION]</c> into a Word section break, so a markdown
/// template can change page orientation — and the page setup with it — partway down a document.
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

        if (ResolvePageMargin(attributes) is { } margins)
        {
            declared.AppendChild(margins);
        }

        properties ??= new();
        properties.SectionProperties = declared;
        renderer.AddBlock(new Paragraph(properties));
        renderer.MayHoldSections = true;
        return true;
    }

    /// <summary>
    /// Rewrite every marker written by <see cref="TryWrite"/> into a real section derived from
    /// <paramref name="baseSection"/>, which is mutated in place to become the document-final section.
    /// </summary>
    /// <remarks>
    /// A paragraph can also carry a <c>w:sectPr</c> this class did not write: one an
    /// <c>OpenXmlToken</c> returned, or one converted html produced. That one is a finished section
    /// rather than a request for one, so it is left exactly as written. It ends the section before
    /// it and says nothing about what follows, so the next marker and the document-final section
    /// carry on from the settings that were in force - which also restates the style source's header
    /// and footer references after it - and, being a section, it counts as one when deciding which
    /// section is the first and may keep a page number start.
    /// </remarks>
    public static void Apply(Body body, SectionProperties? baseSection, string name)
    {
        var sections = body
            .Descendants<Paragraph>()
            .Where(_ => _.ParagraphProperties?.SectionProperties != null)
            .ToList();

        if (sections.Count == 0)
        {
            return;
        }

        if (baseSection == null)
        {
            // Nothing to derive from is only a problem for a marker.
            if (!sections.Any(_ => IsDeclared(_.ParagraphProperties!.SectionProperties!)))
            {
                return;
            }

            throw new ParchmentRenderException(
                name,
                $"'{Marker}' needs a style source to derive its page setup from, and this template has none.");
        }

        // The settings in force at the top of the document. Each marker replaces this with what it
        // declared, so the marker after it — or the document-final section — inherits from there
        // rather than from the style source.
        var current = (SectionProperties) baseSection.CloneNode(true);
        var first = true;

        foreach (var paragraph in sections)
        {
            var properties = paragraph.ParagraphProperties!;
            var declared = properties.SectionProperties!;
            var topLevel = ReferenceEquals(paragraph.Parent, body);

            if (!IsDeclared(declared))
            {
                // Not a section of the body anywhere else, so only a top level one is counted.
                if (topLevel)
                {
                    first = false;
                }

                continue;
            }

            if (!topLevel)
            {
                throw new ParchmentRenderException(
                    name,
                    $"'{Marker}' has to be a paragraph of its own at the top level of the document. " +
                    "One inside a list, a table cell or a block quote is not a section break Word can express.");
            }

            var next = Derive(current, declared);

            // Page numbering restarts wherever w:pgNumType carries a start, so only the first
            // section keeps the style source's. Every later one continues from the section before
            // it, which is what inserting a section break in Word gives you.
            if (!first)
            {
                StripPageNumberStart(current);
            }

            first = false;
            properties.SectionProperties = current;
            current = next;
        }

        if (!first)
        {
            StripPageNumberStart(current);
        }

        // The caller appends baseSection after this returns, so the last declared settings have to
        // land on that instance rather than on a replacement for it.
        baseSection.RemoveAllChildren();
        foreach (var child in current.ChildElements.ToList())
        {
            baseSection.AppendChild(child.CloneNode(true));
        }
    }

    // What TryWrite writes and nothing a finished section looks like: a w:pgSz naming an orientation
    // and no size, with at most a w:pgMar beside it. Told apart by shape rather than by identity
    // because the paragraph may have been cloned or parsed back from xml on its way here.
    static bool IsDeclared(SectionProperties sectPr)
    {
        if (sectPr.GetFirstChild<PageSize>() is not { Orient: not null, Width: null, Height: null })
        {
            return false;
        }

        return sectPr.ChildElements.All(_ => _ is PageSize or PageMargin);
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

        // Asked of every paragraph in the document, and nearly all of them are prose: one that does
        // not open with the marker's bracket is turned away before anything is built. A first
        // literal that is all whitespace says nothing either way, so it falls through.
        if (child is not LiteralInline lead)
        {
            return false;
        }

        var opening = lead.Content.AsSpan().TrimStart();
        if (opening.Length > 0 &&
            opening[0] != '[')
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

    // The page setup a marker can name, all in twips: margins=top,right,bottom,left for the four
    // page edges, in the order css writes them, and header= and footer= for the distances from the
    // top and the bottom of the page to the header and the footer. Those two are their own
    // attributes rather than a fifth and sixth margin because they are a different measurement -
    // taken from the page edge rather than describing where the text sits - and because a template
    // that only wants to move the text should not have to restate them.
    //
    // Each attribute is read independently, and a value that does not parse is ignored rather than
    // rejected, the way an unparseable [TOC] level is: a typo costs that one setting rather than
    // the section break or the settings beside it. Whatever the marker does not name keeps what the
    // style source had.
    static PageMargin? ResolvePageMargin(IEnumerable<KeyValuePair<string, string?>>? attributes)
    {
        if (attributes == null)
        {
            return null;
        }

        PageMargin? margin = null;
        foreach (var (key, value) in attributes)
        {
            if (value == null)
            {
                continue;
            }

            if (string.Equals(key, "margins", StringComparison.OrdinalIgnoreCase))
            {
                if (ParseEdges(value) is { } edges)
                {
                    margin ??= new();
                    margin.Top = edges[0];
                    margin.Right = (uint) edges[1];
                    margin.Bottom = edges[2];
                    margin.Left = (uint) edges[3];
                }
            }
            else if (string.Equals(key, "header", StringComparison.OrdinalIgnoreCase))
            {
                if (ParseDistance(value) is { } header)
                {
                    margin ??= new();
                    margin.Header = header;
                }
            }
            else if (string.Equals(key, "footer", StringComparison.OrdinalIgnoreCase))
            {
                if (ParseDistance(value) is { } footer)
                {
                    margin ??= new();
                    margin.Footer = footer;
                }
            }
        }

        return margin;
    }

    // All four page edges or none of them: a shorter list is as likely to be a mistake as a
    // shorthand, and there is no css shorthand it could be read as here.
    static int[]? ParseEdges(string value)
    {
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

        return twips;
    }

    // Unsigned, where a page edge is not: Word measures the header and footer distances from the
    // page edge inwards and has nowhere to put one that starts outside the page, while a negative
    // margin is a page a header prints into and is a document Word will open.
    static uint? ParseDistance(string value)
    {
        if (uint.TryParse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips))
        {
            return twips;
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
                // Only what the marker named. Anything it left out keeps the style source's, and so
                // does the gutter, which no marker can name: a template changing page setup partway
                // down a document has no opinion on binding.
                target.Top = margins.Top ?? target.Top;
                target.Right = margins.Right ?? target.Right;
                target.Bottom = margins.Bottom ?? target.Bottom;
                target.Left = margins.Left ?? target.Left;
                target.Header = margins.Header ?? target.Header;
                target.Footer = margins.Footer ?? target.Footer;
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
