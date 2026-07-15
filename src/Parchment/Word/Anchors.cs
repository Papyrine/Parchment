/// <summary>
/// Manages the Parchment anchor bookmarks injected at registration time into token-bearing paragraphs.
/// Anchors survive clones intact and let us locate the host paragraph by name at render time, without
/// relying on fragile positional indices that break when structural replacements expand a token into
/// multiple elements.
/// </summary>
static class Anchors
{
    public const string Prefix = "parchment-anchor-";

    // Seed a per-part id counter from MaxBookmarkId once, then pass it by ref so each inserted
    // anchor consumes the next id. The previous per-paragraph NextBookmarkId re-walked the whole
    // part's BookmarkStart descendants on every call — O(paragraphs²) across a token-dense part.
    public static string EnsureOn(Paragraph paragraph, ref int nextId)
    {
        var existing = paragraph
            .Elements<BookmarkStart>()
            .FirstOrDefault(_ => _.Name != null &&
                                 _.Name.Value != null &&
                                 _.Name.Value.StartsWith(Prefix, StringComparison.Ordinal));
        if (existing?.Name?.Value != null)
        {
            return existing.Name.Value;
        }

        var name = Prefix + Guid.NewGuid().ToString("N");
        var idText = (nextId++).ToString();
        var start = new BookmarkStart
        {
            Id = idText,
            Name = name
        };
        var end = new BookmarkEnd
        {
            Id = idText
        };
        InsertAfterProperties(paragraph, start, end);
        return name;
    }

    public static Dictionary<string, Paragraph> BuildMap(OpenXmlCompositeElement root)
    {
        // Parchment-prefixed bookmarks are always direct children of a Paragraph (see
        // InsertAfterProperties). Walk Paragraphs and scan direct children rather than
        // Descendants<BookmarkStart>() over the full subtree.
        var map = new Dictionary<string, Paragraph>(StringComparer.Ordinal);
        foreach (var paragraph in root.Descendants<Paragraph>())
        {
            foreach (var child in paragraph.ChildElements)
            {
                if (child is BookmarkStart {Name.Value: { } name} &&
                    name.StartsWith(Prefix, StringComparison.Ordinal))
                {
                    map[name] = paragraph;
                }
            }
        }

        return map;
    }

    public static void StripAll(OpenXmlCompositeElement root)
    {
        // Single Descendants<Paragraph> pass instead of two Descendants<BookmarkStart> +
        // Descendants<BookmarkEnd> walks over the full subtree. For a fully-expanded body
        // (post-loop), the original two-walk pattern was the dominant non-Save cost in the
        // per-render path (~13% of total at 1000 loop iterations).
        //
        // Parchment bookmarks are inserted Start-then-End in the same paragraph
        // (InsertAfterProperties). Document order therefore guarantees a Start is encountered
        // before its matching End, so the single-pass id-set works correctly.
        List<BookmarkStart>? starts = null;
        List<BookmarkEnd>? ends = null;
        HashSet<string>? ids = null;
        foreach (var paragraph in root.Descendants<Paragraph>())
        {
            foreach (var child in paragraph.ChildElements)
            {
                if (child is BookmarkStart {Name.Value: { } name} start &&
                    name.StartsWith(Prefix, StringComparison.Ordinal))
                {
                    (starts ??= []).Add(start);
                    var id = start.Id?.Value;
                    if (id != null)
                    {
                        (ids ??= new(StringComparer.Ordinal)).Add(id);
                    }
                }
                else if (child is BookmarkEnd {Id.Value: { } endId} end &&
                         ids != null &&
                         ids.Contains(endId))
                {
                    (ends ??= []).Add(end);
                }
            }
        }

        if (starts == null)
        {
            return;
        }

        foreach (var start in starts)
        {
            start.Remove();
        }

        if (ends != null)
        {
            foreach (var end in ends)
            {
                end.Remove();
            }
        }
    }

    /// <summary>
    /// The highest numeric bookmark id already present in <paramref name="root"/>. Callers seed a
    /// per-part id counter from this once (starting at result + 1) and increment locally as anchors
    /// are inserted, rather than re-scanning the part for every anchor. Ids only need to be unique
    /// within the part — Parchment anchors are located by name and stripped before save, so
    /// cross-part id collisions never reach the output.
    /// </summary>
    public static int MaxBookmarkId(OpenXmlCompositeElement root)
    {
        var max = 0;
        foreach (var b in root.Descendants<BookmarkStart>())
        {
            if (int.TryParse(b.Id?.Value, out var value) && value > max)
            {
                max = value;
            }
        }

        return max;
    }

    static void InsertAfterProperties(Paragraph paragraph, BookmarkStart start, BookmarkEnd end)
    {
        var pPr = paragraph.ParagraphProperties;
        if (pPr == null)
        {
            paragraph.InsertAt(end, 0);
            paragraph.InsertAt(start, 0);
        }
        else
        {
            paragraph.InsertAfter(start, pPr);
            paragraph.InsertAfter(end, start);
        }
    }
}
