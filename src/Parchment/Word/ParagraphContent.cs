using M = DocumentFormat.OpenXml.Math;

/// <summary>
/// What a paragraph carries besides its text.
/// </summary>
/// <remarks>
/// <para>
/// A block tag paragraph and the host of a solo structural token are recognised by their text: the
/// tag or the token is all of it. Both are then taken out of the document, and a paragraph can hold
/// things its text does not show — a page break, a picture, a footnote reference. Those are not
/// the tag's to take with it, so the runner asks here first and removes only the text when the
/// answer is yes.
/// </para>
/// <para>
/// Counted: whatever a run holds that is not its properties, its text or white space; an equation
/// written beside the runs; and a page break written as <c>w:pageBreakBefore</c>, which is the
/// same break to whoever wrote it. Not counted, because each would otherwise leave an empty
/// paragraph behind in documents that render cleanly today:
/// </para>
/// <list type="bullet">
/// <item>a tab, a line break, a carriage return or a soft hyphen — white space, as it is in the
/// text;</item>
/// <item><c>w:lastRenderedPageBreak</c> — where Word's last layout happened to end a page, not a
/// break that was asked for;</item>
/// <item>a comment reference and tracked-deleted text — a note about the tag and text that is no
/// longer there;</item>
/// <item>a field — the token is as likely to be the field's result, typed over a merge field, as
/// to sit beside it, and half a field kept in an empty paragraph is worse than none;</item>
/// <item>range markers written beside the runs (bookmarks, comment and permission ranges, proofing
/// marks) — Word parks <c>_GoBack</c> on whichever paragraph was edited last;</item>
/// <item>the rest of the paragraph's formatting, which formats a paragraph that is going away.</item>
/// </list>
/// <para>
/// A section break is not asked about here. <c>w:sectPr</c> sits in the properties of the last
/// paragraph of a section but describes the section, and the runner moves it to whichever
/// paragraph ends the section instead — see <c>ScopeTreeRunner.SettleSectionBreaks</c>.
/// </para>
/// </remarks>
static class ParagraphContent
{
    /// <summary>
    /// Whether <paramref name="paragraph"/> holds content of its own beyond text and white space.
    /// A text box anchored in it counts as a whole; what the box holds is not looked at.
    /// </summary>
    public static bool HasBesidesText(Paragraph paragraph) =>
        HoldsContent(paragraph);

    /// <summary>
    /// Removes the text of <paramref name="paragraph"/> and the runs that held nothing else,
    /// leaving whatever <see cref="HasBesidesText"/> found.
    /// </summary>
    public static void RemoveText(Paragraph paragraph) =>
        RemoveFromRuns(paragraph, _ => _ is Text);

    /// <summary>
    /// Removes the begin, separate and end marks and the instructions of the fields in
    /// <paramref name="paragraph"/>, so that no part of it is left holding half a field.
    /// </summary>
    public static void RemoveFieldMarks(Paragraph paragraph) =>
        RemoveFromRuns(paragraph, _ => _ is FieldChar or FieldCode);

    // Walks by sibling rather than through ParagraphText.Elements: this runs for every tag
    // paragraph a render removes, and nearly all of them hold nothing to find.
    static bool HoldsContent(OpenXmlElement parent)
    {
        for (var child = parent.FirstChild; child != null; child = child.NextSibling())
        {
            // A text box's paragraphs are hosts of their own, as they are to ParagraphText.
            if (child is Paragraph)
            {
                continue;
            }

            if (IsContent(child) ||
                HoldsContent(child))
            {
                return true;
            }
        }

        return false;
    }

    static bool IsContent(OpenXmlElement element)
    {
        if (element is M.OfficeMath or M.Paragraph)
        {
            return true;
        }

        if (element is PageBreakBefore pageBreak)
        {
            return element.Parent is ParagraphProperties &&
                   pageBreak.Val?.Value != false;
        }

        if (element.Parent is not Run)
        {
            return false;
        }

        return element switch
        {
            RunProperties or Text => false,
            TabChar or PositionalTab or CarriageReturn or SoftHyphen => false,
            LastRenderedPageBreak => false,
            CommentReference or DeletedText => false,
            FieldChar or FieldCode or DeletedFieldCode => false,
            Break @break => @break.Type?.Value == BreakValues.Page ||
                            @break.Type?.Value == BreakValues.Column,
            _ => true
        };
    }

    static void RemoveFromRuns(Paragraph paragraph, Func<OpenXmlElement, bool> match)
    {
        foreach (var element in ParagraphText.Elements(paragraph).Where(match).ToList())
        {
            var owner = element.Parent;
            element.Remove();
            if (owner is Run run &&
                run.ChildElements.All(_ => _ is RunProperties))
            {
                run.Remove();
            }
        }
    }
}
