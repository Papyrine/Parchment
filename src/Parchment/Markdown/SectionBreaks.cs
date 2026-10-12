using System.Numerics;

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
/// <para>
/// A markdown value rendered by the docx flow goes through the same walk, so it can hold a marker
/// too, but there the document already has sections of its own and the paragraphs carrying their
/// <c>w:sectPr</c> are not markers. So the runner <see cref="Collect"/>s the markers each value
/// wrote, and <see cref="ApplyInPart"/> derives each from the section of the template it landed in.
/// </para>
/// </remarks>
static class SectionBreaks
{
    const string Marker = "[SECTION]";

    // A4, matching Word's own default, for a style source whose section properties do not say.
    const uint DefaultWidth = 11906;
    const uint DefaultHeight = 16838;

    // The longest side Word gives a page, 22 inches, and so the furthest any measurement on one
    // can reach.
    const int MaxTwips = 31680;

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
            .Where(IsMarker)
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

        Rewrite(body, markers, baseSection, name, partUri: null);
    }

    /// <summary>
    /// The docx flow's <see cref="Apply"/>: rewrite the <paramref name="markers"/> that markdown
    /// values wrote into a part that already has sections of its own.
    /// </summary>
    /// <remarks>
    /// Each run of markers is derived from the <c>w:sectPr</c> that ends the section it landed in -
    /// the next one in document order, on a paragraph of the template's or at the end of the body -
    /// and that <c>w:sectPr</c> is mutated in place to whatever the run's last marker declared. So a
    /// marker still reads "from here on", as far as the template's next section break, where the
    /// page setup the template gave the following section takes over again.
    /// </remarks>
    public static void ApplyInPart(OpenXmlCompositeElement root, List<Paragraph> markers, string name, string partUri)
    {
        if (root is not Body body)
        {
            throw new ParchmentRenderException(
                name,
                $"'{Marker}' is a section break, and only the document body can hold one. " +
                "A header, a footer, a footnote or an endnote cannot.",
                partUri);
        }

        // Every paragraph ending a section, the template's own and the markers alike, listed before
        // anything is rewritten: only a walk of the document says which section each marker is in.
        var written = new HashSet<Paragraph>(markers);
        var breaks = body
            .Descendants<Paragraph>()
            .Where(IsMarker)
            .ToList();

        var pending = new List<Paragraph>();
        foreach (var paragraph in breaks)
        {
            if (written.Contains(paragraph))
            {
                pending.Add(paragraph);
                continue;
            }

            Rewrite(body, pending, paragraph.ParagraphProperties!.SectionProperties!, name, partUri);
            pending.Clear();
        }

        if (pending.Count == 0)
        {
            return;
        }

        if (body.GetFirstChild<SectionProperties>() is not { } final)
        {
            throw new ParchmentRenderException(
                name,
                $"'{Marker}' needs a section to derive its page setup from, and the template's body does not end with a w:sectPr.",
                partUri);
        }

        Rewrite(body, pending, final, name, partUri);
    }

    /// <summary>
    /// Add every marker in <paramref name="produced"/> - what one markdown value rendered to - to
    /// <paramref name="markers"/>. Only for elements <c>MarkdownRendering.Render</c> returned: a
    /// paragraph carrying a <c>w:sectPr</c> is a marker there and nowhere else in the docx flow.
    /// </summary>
    public static void Collect(IEnumerable<OpenXmlElement> produced, List<Paragraph> markers)
    {
        foreach (var element in produced)
        {
            if (element is Paragraph paragraph)
            {
                if (IsMarker(paragraph))
                {
                    markers.Add(paragraph);
                }

                continue;
            }

            // Inside a table the value built. It cannot become a section, and ApplyInPart says so.
            markers.AddRange(element.Descendants<Paragraph>().Where(IsMarker));
        }
    }

    public static bool IsMarker(OpenXmlElement element) =>
        element is Paragraph { ParagraphProperties.SectionProperties: not null };

    // Gives each marker the settings in force before it and leaves what the last one declared on
    // baseSection, which stays where it is.
    static void Rewrite(Body body, List<Paragraph> markers, SectionProperties baseSection, string name, string? partUri)
    {
        if (markers.Count == 0)
        {
            return;
        }

        // The settings in force at the top of the section. Each marker replaces this with what it
        // declared, so the marker after it — or the end of the section — inherits from there
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
                    "One inside a list, a table cell or a block quote is not a section break Word can express.",
                    partUri);
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

        // baseSection is the instance the document keeps - the markdown flow appends it after this
        // returns, the docx flow's is already in place - so the last declared settings have to land
        // on it rather than on a replacement for it.
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
    //
    // A number that parses and cannot be a measurement on a page is another matter, and is rejected:
    // there is no setting to fall back to that the author could have meant, and w:left and w:right
    // are unsigned in the schema, so a negative one would otherwise wrap round to a margin of four
    // billion twips.
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
                    // Top and bottom may be negative, which Word reads as a fixed distance the
                    // header or footer is not allowed to push the text away from. Left and right
                    // have no such reading.
                    margin ??= new();
                    margin.Top = (int) InRange(edges[0], -MaxTwips, "margins", value, "top margin");
                    margin.Right = (uint) InRange(edges[1], 0, "margins", value, "right margin");
                    margin.Bottom = (int) InRange(edges[2], -MaxTwips, "margins", value, "bottom margin");
                    margin.Left = (uint) InRange(edges[3], 0, "margins", value, "left margin");
                }
            }
            else if (string.Equals(key, "header", StringComparison.OrdinalIgnoreCase))
            {
                if (ParseDistance(value) is { } header)
                {
                    margin ??= new();
                    margin.Header = (uint) InRange(header, 0, "header", value, "header distance");
                }
            }
            else if (string.Equals(key, "footer", StringComparison.OrdinalIgnoreCase))
            {
                if (ParseDistance(value) is { } footer)
                {
                    margin ??= new();
                    margin.Footer = (uint) InRange(footer, 0, "footer", value, "footer distance");
                }
            }
        }

        return margin;
    }

    // All four page edges or none of them: a shorter list is as likely to be a mistake as a
    // shorthand, and there is no css shorthand it could be read as here.
    //
    // Read without a width, so a number too large for an int is still a number - one InRange turns
    // away by name - rather than something that failed to parse and was quietly dropped.
    static BigInteger[]? ParseEdges(string value)
    {
        var parts = value.Split(',');
        if (parts.Length != 4)
        {
            return null;
        }

        var twips = new BigInteger[4];
        for (var index = 0; index < 4; index++)
        {
            if (!BigInteger.TryParse(parts[index].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out twips[index]))
            {
                return null;
            }
        }

        return twips;
    }

    // Unsigned, where a page edge is not: Word measures the header and footer distances from the
    // page edge inwards and has nowhere to put one that starts outside the page, while a negative
    // margin is a page a header prints into and is a document Word will open.
    static BigInteger? ParseDistance(string value)
    {
        if (BigInteger.TryParse(value.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips) &&
            twips.Sign >= 0)
        {
            return twips;
        }

        return null;
    }

    static BigInteger InRange(BigInteger twips, int min, string attribute, string value, string what)
    {
        if (twips >= min &&
            twips <= MaxTwips)
        {
            return twips;
        }

        throw new SectionMarkerException(
            $"'{Marker}' has {attribute}={value}, whose {what} of {twips} twips is out of range. " +
            $"It has to be between {min} and {MaxTwips}.");
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
                // By the schema's order rather than at the end: w:sectPr is a strict sequence, and a
                // style source without margins can still carry the w:cols or w:docGrid that follow.
                next.AddChild(margins.CloneNode(true));
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

        // By the schema's order, since nearly everything else in a w:sectPr has to follow it.
        sectPr.AddChild(size);
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
