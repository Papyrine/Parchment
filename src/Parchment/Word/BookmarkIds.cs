/// <summary>
/// Per-render allocator for <c>w:bookmarkStart</c>/<c>w:bookmarkEnd</c> ids, which have to be
/// unique across the whole document — headers, footers and notes included.
/// </summary>
/// <remarks>
/// <para>
/// Nothing that produces content during a render can see what ids are already taken. The html
/// converter restarts at 1 for every conversion, the markdown renderer for every render, and an
/// <see cref="OpenXmlToken"/> is caller code. So every produced fragment is renumbered here as it
/// is produced, from one running counter, and a generated <c>_Toc</c> bookmark draws from the same.
/// Only the fragment is walked, never the document it lands in. Renumbering is safe because nothing
/// refers to a bookmark by id: a link's <c>w:anchor</c> and a <c>PAGEREF</c> both name it.
/// </para>
/// <para>
/// The counter is seeded at registration with the highest id the template carries, the anchor
/// bookmarks included. Anchors are gone before the save, but <see cref="Anchors.StripAll"/> pairs
/// an anchor's end to its start by id, so a produced bookmark that shared an anchor's id would lose
/// its end to that pass. Starting above them leaves a gap in the output's ids, which costs nothing:
/// the ids have to be distinct, not dense. Sequential in processing order, so the same template and
/// model give the same ids.
/// </para>
/// </remarks>
sealed class BookmarkIds(int max)
{
    int last = max;

    // Beside the ids because both are per render and both reach everything that writes a bookmark.
    public BookmarkNames Names { get; } = new();

    public string Next() =>
        (++last).ToString(CultureInfo.InvariantCulture);

    public void Renumber(OpenXmlElement element)
    {
        Dictionary<string, string>? open = null;
        Renumber(element, ref open);
    }

    /// <summary>
    /// Give every bookmark in <paramref name="elements"/> a fresh id, keeping each end on the id of
    /// the start it closes. The elements are taken as one fragment, since a bookmark can open in one
    /// block and close in a later one.
    /// </summary>
    public void Renumber(IReadOnlyList<OpenXmlElement> elements)
    {
        Dictionary<string, string>? open = null;
        foreach (var element in elements)
        {
            Renumber(element, ref open);
        }
    }

    void Renumber(OpenXmlElement element, ref Dictionary<string, string>? open)
    {
        foreach (var bookmark in Bookmarks(element))
        {
            if (bookmark is BookmarkStart start)
            {
                var id = start.Id?.Value;
                if (id == null)
                {
                    continue;
                }

                var replacement = Next();
                open ??= new(StringComparer.Ordinal);
                open[id] = replacement;
                start.Id = replacement;
            }
            else if (bookmark is BookmarkEnd end)
            {
                var id = end.Id?.Value;
                if (id != null &&
                    open != null &&
                    open.Remove(id, out var replacement))
                {
                    end.Id = replacement;
                }
            }
        }
    }

    // Bookmark markers in document order, the element itself included — a bookmark can be a block's
    // own sibling as well as something nested inside it.
    static IEnumerable<OpenXmlElement> Bookmarks(OpenXmlElement element)
    {
        if (element is BookmarkStart or BookmarkEnd)
        {
            yield return element;
            yield break;
        }

        foreach (var descendant in element.Descendants<OpenXmlElement>())
        {
            if (descendant is BookmarkStart or BookmarkEnd)
            {
                yield return descendant;
            }
        }
    }
}
