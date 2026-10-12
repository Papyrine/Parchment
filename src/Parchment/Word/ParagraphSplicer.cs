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
/// token becomes another paragraph (cloning the host's <c>pPr</c>).
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
    public static void SpliceInline(Paragraph host, int offset, int length, Paragraph producedParagraph)
    {
        var beforeChildren = TrimmedHead(host, offset);
        var afterChildren = TrimmedTail(host, offset, offset + length);
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
        IReadOnlyList<OpenXmlElement> produced)
    {
        var after = BuildShellWithChildren(host, TrimmedTail(host, offset, offset + length));
        TrimToHead(host, offset);

        var result = new List<OpenXmlElement>(produced.Count + 1);
        foreach (var element in produced)
        {
            result.Add(element.CloneNode(true));
        }

        result.Add(after);
        return result;
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
    /// and removed when it sits past the offset. Emptied runs are then removed.
    /// </summary>
    static void TrimToHead(Paragraph paragraph, int upToOffset)
    {
        var consumed = 0;
        var toRemove = new List<OpenXmlElement>();

        foreach (var element in ParagraphText.Elements(paragraph).ToList())
        {
            if (element is not Text text)
            {
                if (IsRunContent(element) &&
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
    /// <paramref name="fromOffset"/>, removes or trims it. Run content that carries no text is
    /// removed when it sits at or before <paramref name="tokenOffset"/> — exactly the content
    /// <see cref="TrimToHead"/> keeps, so each such element survives in one half only. Emptied
    /// runs are then removed.
    /// </summary>
    static void TrimToTail(Paragraph paragraph, int tokenOffset, int fromOffset)
    {
        var consumed = 0;
        var toRemove = new List<OpenXmlElement>();

        foreach (var element in ParagraphText.Elements(paragraph).ToList())
        {
            if (element is not Text text)
            {
                if (IsRunContent(element) &&
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
