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
/// A <see cref="MutateToken"/> is caller code as well, but it produces no fragment: it edits its
/// host paragraph in place, among markers that are the template's and the anchor's. So the markers
/// the host holds are noted before the delegate runs (<see cref="Markers"/>) and what it added
/// beyond them is renumbered afterwards (<see cref="RenumberAdded"/>). The host is all that is
/// walked: a marker the delegate puts outside its own paragraph keeps the id it was given.
/// </para>
/// <para>
/// The counter is seeded at registration with the highest id the template carries, the anchor
/// bookmarks included, so nothing numbered here shares an id with an anchor while both are in the
/// document. Anchors are gone before the save, which leaves a gap in the output's ids and costs
/// nothing: the ids have to be distinct, not dense. Sequential in processing order, so the same
/// template and model give the same ids.
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

    /// <summary>
    /// The bookmark markers <paramref name="element"/> holds, taken before caller code edits it in
    /// place so that <see cref="RenumberAdded"/> can tell what the caller added.
    /// </summary>
    public static HashSet<OpenXmlElement> Markers(OpenXmlElement element)
    {
        var markers = new HashSet<OpenXmlElement>(ReferenceEqualityComparer.Instance);
        foreach (var bookmark in Bookmarks(element))
        {
            markers.Add(bookmark);
        }

        return markers;
    }

    /// <summary>
    /// Give every bookmark put into <paramref name="element"/> since <paramref name="existing"/>
    /// was taken a fresh id. Only one that both opens and closes inside the element: a bookmark
    /// that reaches outside it is left on the id it was given, since renumbering the half that is
    /// in view would part it from the half that is not.
    /// </summary>
    public void RenumberAdded(OpenXmlElement element, HashSet<OpenXmlElement> existing)
    {
        Dictionary<string, BookmarkStart>? open = null;
        foreach (var bookmark in Bookmarks(element))
        {
            if (existing.Contains(bookmark))
            {
                continue;
            }

            if (bookmark is BookmarkStart start)
            {
                var id = start.Id?.Value;
                if (id != null)
                {
                    open ??= new(StringComparer.Ordinal);
                    open[id] = start;
                }
            }
            else if (bookmark is BookmarkEnd end)
            {
                var id = end.Id?.Value;
                if (id != null &&
                    open != null &&
                    open.Remove(id, out var opened))
                {
                    var replacement = Next();
                    opened.Id = replacement;
                    end.Id = replacement;
                }
            }
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
