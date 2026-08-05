/// <summary>
/// Hands out bookmarks for paragraphs that need to be referable, reusing one the paragraph already
/// carries rather than adding a second name for the same place.
/// </summary>
/// <remarks>
/// A table-of-contents entry has to link somewhere, and most headings were never given an id by the
/// template author. Generated names follow Word's own <c>_Toc…</c> convention, which is what Word
/// writes when it builds a table of contents and what it recognises as one of its own.
/// </remarks>
class WordBookmarkAllocator
{
    int next;

    public WordBookmarkAllocator(Body body) =>
        next = body.Descendants<BookmarkStart>()
            .Select(_ => int.TryParse(_.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : 0)
            .DefaultIfEmpty(0)
            .Max();

    public string Ensure(Paragraph paragraph)
    {
        var existing = paragraph.Descendants<BookmarkStart>().FirstOrDefault()?.Name?.Value;
        if (existing != null)
        {
            return existing;
        }

        var id = (++next).ToString(CultureInfo.InvariantCulture);
        var name = $"_Toc{id}";
        var start = new BookmarkStart
        {
            Id = id,
            Name = name
        };

        // Inside the paragraph and around its runs, so a page reference to it resolves to the page
        // the heading is actually on rather than wherever an empty marker happened to fall.
        var properties = paragraph.ParagraphProperties;
        if (properties == null)
        {
            paragraph.InsertAt(start, 0);
        }
        else
        {
            properties.InsertAfterSelf(start);
        }

        paragraph.Append(
            new BookmarkEnd
            {
                Id = id
            });
        return name;
    }
}
