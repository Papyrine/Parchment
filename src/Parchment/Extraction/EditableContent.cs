namespace Parchment;

/// <summary>
/// The one definition of what an editable control holds, shared by the plain-text reader
/// (<see cref="EditableFieldReader.RawText"/>) and the rich-text one
/// (<see cref="OpenXmlHtmlSerializer"/>) so the two cannot disagree about which text counts.
/// The content is read as Word shows it with every tracked change accepted: insertions
/// (<c>w:ins</c>) and move destinations (<c>w:moveTo</c>) are kept, deletions (<c>w:del</c>) and
/// move sources (<c>w:moveFrom</c>) are skipped. Every other wrapper is transparent — simple
/// fields (the result runs; the instruction is an attribute), nested content controls,
/// <c>w:smartTag</c>, <c>w:customXml</c>, <c>w:dir</c>, <c>w:bdo</c>, tables, and anything not
/// known here — because the walk descends by default and names only what it skips.
/// </summary>
static class EditableContent
{
    const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    const string compatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>
    /// The paragraphs of a block container in document order, through nested content controls
    /// and table cells. A paragraph inside a text box is not one of them: a text box is anchored
    /// in a run, so its text arrives through <see cref="Inline"/>.
    /// </summary>
    public static IEnumerable<Paragraph> Paragraphs(OpenXmlElement container)
    {
        foreach (var child in Visible(container))
        {
            if (child is Paragraph paragraph)
            {
                if (!IsRemoved(paragraph))
                {
                    yield return paragraph;
                }

                continue;
            }

            foreach (var nested in Paragraphs(child))
            {
                yield return nested;
            }
        }
    }

    /// <summary>
    /// The inline pieces of a container in document order: each <see cref="Run"/>, each
    /// <see cref="Hyperlink"/> (not entered — pass it back in for its own pieces), and each
    /// <c>m:oMath</c>, which has no runs of its own and is read as its text. The runs of a text
    /// box follow the run that anchors it.
    /// </summary>
    public static IEnumerable<OpenXmlElement> Inline(OpenXmlElement container)
    {
        foreach (var child in Visible(container))
        {
            if (child is Hyperlink ||
                IsMath(child))
            {
                yield return child;
                continue;
            }

            if (child is Run)
            {
                yield return child;
            }

            foreach (var nested in Inline(child))
            {
                yield return nested;
            }
        }
    }

    /// <summary>
    /// <see cref="Inline"/> with every hyperlink replaced by its own pieces.
    /// </summary>
    public static IEnumerable<OpenXmlElement> Runs(OpenXmlElement container)
    {
        foreach (var piece in Inline(container))
        {
            if (piece is not Hyperlink)
            {
                yield return piece;
                continue;
            }

            foreach (var nested in Runs(piece))
            {
                yield return nested;
            }
        }
    }

    // The children a reader of the accepted document sees. Typed checks would miss the SDK's
    // several classes per tracked-change element (w:del is one class in a paragraph and another
    // in a table row's properties), so these go by name.
    static IEnumerable<OpenXmlElement> Visible(OpenXmlElement container)
    {
        foreach (var child in container.Elements())
        {
            if (child.NamespaceUri == compatibilityNamespace)
            {
                // mc:AlternateContent states one piece of content several ways (a text box is a
                // DrawingML shape and a VML fallback). Read one of them, the preferred one.
                if (child.LocalName == "AlternateContent" &&
                    child.Elements().FirstOrDefault() is { } preferred)
                {
                    foreach (var nested in Visible(preferred))
                    {
                        yield return nested;
                    }
                }

                continue;
            }

            if (child.LocalName.EndsWith("Pr", StringComparison.Ordinal))
            {
                continue;
            }

            if (child.NamespaceUri == wordNamespace &&
                child.LocalName is "del" or "moveFrom")
            {
                continue;
            }

            if (child is TableRow row &&
                HasRemovalMark(row.TableRowProperties))
            {
                continue;
            }

            yield return Typed(child);
        }
    }

    // Below an element the SDK has no class for (w:smartTag, the shape around a text box) every
    // descendant loads as an unknown element, a w:r included. Re-read the three kinds the readers
    // match on as what they are.
    static OpenXmlElement Typed(OpenXmlElement element)
    {
        if (element is not OpenXmlUnknownElement ||
            element.NamespaceUri != wordNamespace)
        {
            return element;
        }

        return element.LocalName switch
        {
            "p" => new Paragraph(element.OuterXml),
            "r" => new Run(element.OuterXml),
            "hyperlink" => new Hyperlink(element.OuterXml),
            _ => element
        };
    }

    // A paragraph whose mark is deleted or moved away, with nothing left in it, is gone once the
    // change is accepted. One with surviving text stays a paragraph of its own.
    static bool IsRemoved(Paragraph paragraph) =>
        HasRemovalMark(paragraph.ParagraphProperties?.ParagraphMarkRunProperties) &&
        !Inline(paragraph).Any();

    static bool HasRemovalMark(OpenXmlElement? properties) =>
        properties != null &&
        properties.Elements().Any(_ => _.LocalName is "del" or "moveFrom");

    static bool IsMath(OpenXmlElement element) =>
        element.NamespaceUri == mathNamespace &&
        element.LocalName == "oMath";
}
