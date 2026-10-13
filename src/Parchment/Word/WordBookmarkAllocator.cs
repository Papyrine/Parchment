/// <summary>
/// Hands out bookmarks for paragraphs that need to be referable, reusing one the paragraph already
/// carries rather than adding a second name for the same place.
/// </summary>
/// <remarks>
/// A table-of-contents entry has to link somewhere, and most headings were never given an id by the
/// template author. Generated names follow Word's own <c>_Toc…</c> convention, which is what Word
/// writes when it builds a table of contents and what it recognises as one of its own.
/// Ids come from the render's <see cref="BookmarkIds"/>, which everything else in the document was
/// numbered from, so no walk of the body is needed to find a free one.
/// </remarks>
class WordBookmarkAllocator(BookmarkIds ids)
{
    public string Ensure(Paragraph paragraph)
    {
        var existing = paragraph.Descendants<BookmarkStart>().FirstOrDefault()?.Name?.Value;
        if (existing != null)
        {
            return existing;
        }

        // An author can write {#_Toc7} too, so a generated name is claimed like any other and the
        // next id tried when it is taken.
        string id;
        string name;
        do
        {
            id = ids.Next();
            name = $"_Toc{id}";
        } while (!ids.Names.Reserve(name));

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
