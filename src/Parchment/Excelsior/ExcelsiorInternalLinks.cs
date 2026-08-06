/// <summary>
/// Turns the <c>#name</c> links in a built Excelsior table into the two things a <c>#name</c> means
/// inside a Word document: a jump to a bookmark, or — with nothing to click — the page that bookmark
/// is on.
/// </summary>
/// <remarks>
/// Excelsior registers every <c>Link</c> as an external relationship, which is right for a
/// spreadsheet and for <c>https://</c>, but a <c>#name</c> target addresses somewhere in this
/// document. Registered externally it is a link Word will not follow, so the rewrite happens here
/// rather than in Excelsior: Parchment owns the table once it is built, and the same two meanings
/// are already what <c>LinkInlineRenderer</c> gives a markdown <c>[text](#name)</c>. A table's links
/// behaving differently from the prose around it would be the surprising part.
/// </remarks>
static class ExcelsiorInternalLinks
{
    // What Word itself caches for a page reference it has not resolved yet, and what
    // LinkInlineRenderer uses for the same field.
    const string pageReferencePlaceholder = "1";

    public static void Rewrite(Table table, MainDocumentPart mainPart)
    {
        // Materialized before mutating: an empty link is replaced by the runs of a field, which
        // edits the tree being walked.
        foreach (var hyperlink in table.Descendants<Hyperlink>().ToList())
        {
            if (hyperlink.Id?.Value is not { } relationshipId)
            {
                continue;
            }

            var relationship = mainPart.HyperlinkRelationships
                .FirstOrDefault(_ => _.Id == relationshipId);
            if (relationship?.Uri.OriginalString is not { } target ||
                target.Length == 0 ||
                target[0] != '#')
            {
                continue;
            }

            var bookmark = MarkdownBookmark.Sanitize(target[1..]);

            // The relationship goes either way: an anchored hyperlink does not carry one, and a
            // field is not a hyperlink at all. Left behind it is a dangling external reference to
            // a fragment, which is what Word warns about when opening the file.
            mainPart.DeleteReferenceRelationship(relationshipId);
            hyperlink.Id = null;

            if (hyperlink.InnerText.Length == 0)
            {
                ReplaceWithPageReference(hyperlink, bookmark);
                continue;
            }

            hyperlink.Anchor = bookmark;
        }
    }

    static void ReplaceWithPageReference(Hyperlink hyperlink, string bookmark)
    {
        var parent = hyperlink.Parent;
        if (parent == null)
        {
            return;
        }

        OpenXmlElement previous = hyperlink;
        foreach (var run in WordFields.PageReference(bookmark, pageReferencePlaceholder))
        {
            previous = parent.InsertAfter(run, previous);
        }

        hyperlink.Remove();
    }
}
