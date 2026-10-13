/// <summary>
/// Finds the block a table would sit directly beside, for the places rendered markdown can leave
/// two tables together.
/// </summary>
/// <remarks>
/// <para>
/// Word reads two <c>w:tbl</c> siblings with nothing between them as one table: the rows of the
/// second continue the first, whatever its own columns were. The schema allows the pair and
/// <c>OpenXmlValidator</c> passes it, so every place that can produce one puts an empty paragraph
/// between them instead - the same separator <c>ScopeTreeRunner.SeparateTables</c> writes
/// for the docx flow.
/// </para>
/// <para>
/// A pipe table directly followed by an <c>[ExcelsiorTable]</c> was seen to stay two tables in
/// Word with nothing between them. What sets that pair apart is the table style - one carries no
/// <c>w:tblStyle</c>, the other Excelsior's - and Word appears to merge only tables whose styles
/// match. That is an accident of the two defaults, gone as soon as the pipe table is given the same
/// style, so nothing here reads the style and that pair is separated like any other.
/// </para>
/// </remarks>
static class AdjacentTables
{
    /// <summary>
    /// Whether a table appended to <paramref name="blocks"/> would directly follow another one.
    /// </summary>
    public static bool EndsWithTable(List<OpenXmlElement> blocks)
    {
        for (var index = blocks.Count - 1; index >= 0; index--)
        {
            if (!IsRangeMarker(blocks[index]))
            {
                return blocks[index] is Table;
            }
        }

        return false;
    }

    public static OpenXmlElement? BlockBefore(OpenXmlElement element)
    {
        var cursor = element.PreviousSibling();
        while (cursor != null &&
               IsRangeMarker(cursor))
        {
            cursor = cursor.PreviousSibling();
        }

        return cursor;
    }

    public static OpenXmlElement? BlockAfter(OpenXmlElement element)
    {
        var cursor = element.NextSibling();
        while (cursor != null &&
               IsRangeMarker(cursor))
        {
            cursor = cursor.NextSibling();
        }

        return cursor;
    }

    // Range markers can sit between two blocks without keeping their tables apart.
    static bool IsRangeMarker(OpenXmlElement element) =>
        element is BookmarkStart or BookmarkEnd or PermStart or PermEnd or ProofError;
}
