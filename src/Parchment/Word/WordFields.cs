/// <summary>
/// Builds the run sequences for the Word field codes markdown can ask for — a table of contents and
/// a page reference.
/// </summary>
/// <remarks>
/// A field is three or five runs rather than one element: <c>begin</c>, the instruction text, an
/// optional <c>separate</c> plus cached result, and <c>end</c>. The result is what Word displays
/// until it recalculates, so a placeholder goes there rather than nothing — an empty result renders
/// as a blank line in readers that never update fields.
/// <para>
/// Every field is marked dirty, which is what makes Word offer to update it (and update it outright
/// for a TOC) when the document opens. Page numbers cannot be computed here: they are a product of
/// layout, which only a renderer knows, so the field carries the instruction and Word supplies the
/// number.
/// </para>
/// </remarks>
static class WordFields
{
    public static IEnumerable<OpenXmlElement> TableOfContents(int maxLevel, string placeholder) =>
        Build($" TOC \\o \"1-{maxLevel}\" \\h \\z \\u ", placeholder);

    public static IEnumerable<OpenXmlElement> PageReference(string bookmarkName, string placeholder) =>
        Build($" PAGEREF {bookmarkName} \\h ", placeholder);

    static IEnumerable<OpenXmlElement> Build(string instruction, string placeholder)
    {
        yield return new Run(
            new FieldChar
            {
                FieldCharType = FieldCharValues.Begin,
                Dirty = true
            });
        yield return new Run(
            new FieldCode(instruction)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        yield return new Run(
            new FieldChar
            {
                FieldCharType = FieldCharValues.Separate
            });
        yield return new Run(
            new Text(placeholder)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        yield return new Run(
            new FieldChar
            {
                FieldCharType = FieldCharValues.End
            });
    }
}
