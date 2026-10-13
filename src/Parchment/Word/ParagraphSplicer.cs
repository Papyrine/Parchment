/// <summary>
/// Splices structurally-rendered content (HTML, markdown) into a token-bearing paragraph when
/// the token does not occupy the entire paragraph.
///
/// Two modes are supported:
///
/// <list type="bullet">
/// <item>
/// <c>Inline splice</c> — the produced output is a single paragraph (typical for inline-only HTML
/// like <c>&lt;b&gt;x&lt;/b&gt;</c> or single-line markdown). The produced paragraph's children
/// are extracted (its <c>pPr</c> is dropped) and spliced into the host paragraph at the token
/// offset, replacing the token text. The host's run-level surroundings and paragraph properties
/// are preserved.
/// </item>
/// <item>
/// <c>Split</c> — the produced output is multiple block elements, or contains a non-paragraph
/// block (a table). The host paragraph is split at the token offset: the host is trimmed to the
/// text before the token, the produced block elements are inserted after it, and text after the
/// token becomes another paragraph (cloning the host's <c>pPr</c>, apart from the properties
/// that belong to one end of the paragraph — see <see cref="DivideProperties"/>).
/// </item>
/// </list>
///
/// Solo tokens (token covers the entire paragraph text) bypass this splicer — see the simpler
/// "swap host paragraph" path in <see cref="ScopeTreeRunner"/>. That path splits the host here only
/// when it carries something besides the token, to keep it on the side of the token it sat on.
/// </summary>
static class ParagraphSplicer
{
    /// <summary>
    /// True when the produced element list is a single Paragraph and so can be unwrapped and
    /// spliced inline. Anything else (multiple blocks, a table) requires host splitting.
    /// </summary>
    public static bool IsInlineEquivalent(IReadOnlyList<OpenXmlElement> produced) =>
        produced is [Paragraph];

    /// <summary>
    /// Inline splice: rebuilds host's children as
    /// <c>[host's children before token] + [produced paragraph's children, minus pPr] + [host's children after token]</c>.
    /// Mutates host in place; no other paragraphs are added or removed.
    /// </summary>
    public static void SpliceInline(Paragraph host, int offset, int length, Paragraph producedParagraph, CloneIds? cloneIds)
    {
        var beforeChildren = TrimmedHead(host, offset);
        var afterChildren = TrimmedTail(host, offset, offset + length);

        // Both are copies of the host, so a wrapper the token sat inside is in each of them.
        cloneIds?.RenumberShared(beforeChildren, afterChildren);
        var producedChildren = ContentChildren(producedParagraph);

        var pPr = host.ParagraphProperties;

        // Drop every non-pPr child from host, then re-add in the new order.
        foreach (var child in host.ChildElements.ToList())
        {
            if (child is ParagraphProperties)
            {
                continue;
            }

            child.Remove();
        }

        OpenXmlElement? cursor = pPr;
        foreach (var element in beforeChildren.Concat(producedChildren).Concat(afterChildren))
        {
            cursor = cursor == null
                ? host.InsertAt(element, 0)
                : host.InsertAfter(element, cursor);
        }
    }

    /// <summary>
    /// Split: trims host in place to the text before the token and returns the ordered list to
    /// insert after it (the caller is responsible for inserting these into host's parent). The
    /// list is <c>[...produced, after-paragraph]</c>.
    /// Host stays the before half rather than being cloned so that tokens at lower offsets, which
    /// are substituted after the split, still land in the document with their offsets intact.
    /// Empty before/after paragraphs are still emitted — the user authored a paragraph there
    /// and removing it would shift the document layout.
    /// </summary>
    public static IReadOnlyList<OpenXmlElement> Split(
        Paragraph host,
        int offset,
        int length,
        IReadOnlyList<OpenXmlElement> produced,
        CloneIds? cloneIds)
    {
        var after = BuildShellWithChildren(host, TrimmedTail(host, offset, offset + length));
        TrimToHead(host, offset);
        DivideProperties(host, after);

        // The after half is a copy of the host: its paragraph properties, and any wrapper the
        // token sat inside, are now in both.
        cloneIds?.RenumberShared([host], [after]);

        var result = new List<OpenXmlElement>(produced.Count + 1);
        foreach (var element in produced)
        {
            result.Add(element.CloneNode(true));
        }

        result.Add(after);
        return result;
    }

    /// <summary>
    /// Both halves of a split start out with the host's paragraph properties. That is right for
    /// formatting and wrong for the properties that mark one end of the paragraph, which have to
    /// survive on one half only. A section break closes its section after the paragraph that
    /// carries it, so it goes to the trailing half: that one is emitted even when empty, which
    /// also gives the break a paragraph to sit on when the produced content ends with a table. A
    /// page break before and a list number open the paragraph, so they stay on the leading half.
    /// Only what the paragraph states itself is divided; the same settings inherited from its
    /// style still reach both halves.
    /// </summary>
    static void DivideProperties(Paragraph before, Paragraph after)
    {
        if (before.ParagraphProperties is { } leading)
        {
            Drop(leading, leading.SectionProperties);
        }

        if (after.ParagraphProperties is not { } trailing)
        {
            return;
        }

        // Number 0 is "not numbered". Removing the element instead would hand the paragraph
        // back to whatever numbering its style carries.
        if (trailing.NumberingProperties is { } numbering)
        {
            numbering.NumberingId = new()
            {
                Val = 0
            };
        }

        // Written as off, the element cancels a break the style asks for, and still has to.
        if (trailing.PageBreakBefore is { } pageBreak &&
            (pageBreak.Val?.Value ?? true))
        {
            Drop(trailing, pageBreak);
        }
    }

    static void Drop(ParagraphProperties properties, OpenXmlElement? element)
    {
        if (element == null)
        {
            return;
        }

        element.Remove();
        if (!properties.HasChildren)
        {
            properties.Remove();
        }
    }

    /// <summary>
    /// Returns deep-cloned non-pPr children of the source paragraph.
    /// </summary>
    static List<OpenXmlElement> ContentChildren(Paragraph source) =>
        source.ChildElements
            .Where(_ => _ is not ParagraphProperties)
            .Select(_ => _.CloneNode(true))
            .ToList();

    /// <summary>
    /// Builds a fresh paragraph that has the same paragraph properties as <paramref name="shell"/>
    /// but whose body is exactly <paramref name="children"/>.
    /// </summary>
    static Paragraph BuildShellWithChildren(Paragraph shell, IReadOnlyList<OpenXmlElement> children)
    {
        var result = new Paragraph();
        if (shell.ParagraphProperties is { } pPr)
        {
            result.AppendChild((ParagraphProperties)pPr.CloneNode(true));
        }

        foreach (var child in children)
        {
            result.AppendChild(child);
        }

        return result;
    }

    /// <summary>
    /// Returns deep-cloned children of host (excluding pPr) trimmed to keep only inner text up to
    /// (but not including) <paramref name="upToOffset"/>.
    /// </summary>
    static List<OpenXmlElement> TrimmedHead(Paragraph host, int upToOffset)
    {
        var clone = (Paragraph)host.CloneNode(true);
        TrimToHead(clone, upToOffset);
        return ContentChildren(clone);
    }

    /// <summary>
    /// Returns deep-cloned children of host (excluding pPr) trimmed to keep only inner text from
    /// <paramref name="fromOffset"/> onward. <paramref name="tokenOffset"/> is where the token
    /// being replaced starts — the boundary <see cref="TrimmedHead"/> was trimmed to.
    /// </summary>
    static List<OpenXmlElement> TrimmedTail(Paragraph host, int tokenOffset, int fromOffset)
    {
        var clone = (Paragraph)host.CloneNode(true);
        TrimToTail(clone, tokenOffset, fromOffset);
        return ContentChildren(clone);
    }

    /// <summary>
    /// Mutates the paragraph: for every Text descendant whose character range falls fully or
    /// partly past <paramref name="upToOffset"/>, removes or trims it. Run content that carries
    /// no text (a tab, a break, a drawing, a field char) is positioned by the text preceding it
    /// and removed when it sits past the offset, and so is either end of a comment range. Emptied
    /// runs are then removed.
    /// </summary>
    static void TrimToHead(Paragraph paragraph, int upToOffset)
    {
        var consumed = 0;
        var toRemove = new List<OpenXmlElement>();

        foreach (var element in ParagraphText.Elements(paragraph).ToList())
        {
            if (element is not Text text)
            {
                if (IsPlaced(element) &&
                    consumed > upToOffset)
                {
                    toRemove.Add(element);
                }

                continue;
            }

            var value = text.Text;
            var start = consumed;
            var end = consumed + value.Length;
            consumed = end;

            if (start >= upToOffset)
            {
                toRemove.Add(text);
                continue;
            }

            if (end > upToOffset)
            {
                var local = upToOffset - start;
                text.Text = value[..local];
                if (text.Text.Length > 0)
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        RemoveContentAndEmptyRuns(toRemove);
    }

    /// <summary>
    /// Mutates the paragraph: for every Text descendant whose character range ends at or before
    /// <paramref name="fromOffset"/>, removes or trims it. Run content that carries no text, or
    /// either end of a comment range, is removed when it sits at or before
    /// <paramref name="tokenOffset"/> — exactly the content <see cref="TrimToHead"/> keeps, so each
    /// such element survives in one half only. Emptied runs are then removed.
    /// </summary>
    static void TrimToTail(Paragraph paragraph, int tokenOffset, int fromOffset)
    {
        var consumed = 0;
        var toRemove = new List<OpenXmlElement>();

        foreach (var element in ParagraphText.Elements(paragraph).ToList())
        {
            if (element is not Text text)
            {
                if (IsPlaced(element) &&
                    consumed <= tokenOffset)
                {
                    toRemove.Add(element);
                }

                continue;
            }

            var value = text.Text;
            var start = consumed;
            var end = consumed + value.Length;
            consumed = end;

            if (end <= fromOffset)
            {
                toRemove.Add(text);
                continue;
            }

            if (start < fromOffset)
            {
                var local = fromOffset - start;
                text.Text = value[local..];
                if (text.Text.Length > 0)
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        RemoveContentAndEmptyRuns(toRemove);
    }

    /// <summary>
    /// True for a run child other than its <c>rPr</c> and its text: a tab, a break, a drawing, a
    /// field char, a symbol, a footnote reference. Paragraph-level zero-width markers (bookmarks)
    /// are not run content and stay in both halves.
    /// </summary>
    static bool IsRunContent(OpenXmlElement element) =>
        element.Parent is Run &&
        element is not RunProperties;

    /// <summary>
    /// True for what goes to one half only, by where it sits: run content, and either end of a
    /// comment range. A comment has one range, so unlike a bookmark its ends cannot be left in
    /// both halves — the start stays with the text before the token and the end with the text
    /// after it, which keeps what the token produced inside the range the author drew.
    /// </summary>
    static bool IsPlaced(OpenXmlElement element) =>
        IsRunContent(element) ||
        element is CommentRangeStart or CommentRangeEnd;

    static void RemoveContentAndEmptyRuns(IReadOnlyList<OpenXmlElement> elements)
    {
        var owners = new HashSet<Run>();
        foreach (var element in elements)
        {
            if (element.Parent is Run run)
            {
                owners.Add(run);
            }

            element.Remove();
        }

        foreach (var run in owners)
        {
            if (run.ChildElements.All(_ => _ is RunProperties))
            {
                run.Remove();
            }
        }
    }
}
