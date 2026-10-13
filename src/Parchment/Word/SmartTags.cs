/// <summary>
/// Gives the content of a <c>w:smartTag</c> its typed elements back.
/// </summary>
/// <remarks>
/// Older versions of Word wrap a recognised name, date or place in a smart tag without the author
/// asking, so a token can sit inside one unnoticed. The Open XML SDK has no class for the element:
/// it loads as an <see cref="OpenXmlUnknownElement"/>, and so does everything inside it, which
/// leaves the runs and text of a token invisible to every walk that looks for a <see cref="Run"/>
/// or a <see cref="Text"/>. A smart tag holds what a paragraph holds, so each child is read again
/// as a paragraph's child and swapped in. The smart tag itself stays, as the unknown element it
/// was, and is written back out around its content.
///
/// Nothing here survives a save: the part is parsed afresh by every render, and the smart tag is
/// unknown again. <see cref="ParagraphText.Elements"/> therefore calls this as it reaches one,
/// which costs a paragraph without a smart tag one type test per unknown element it holds.
/// </remarks>
static class SmartTags
{
    const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void Retype(OpenXmlUnknownElement element)
    {
        if (element.LocalName != "smartTag" ||
            element.NamespaceUri != wordNamespace)
        {
            return;
        }

        var child = element.FirstChild;
        while (child != null)
        {
            var next = child.NextSibling();
            if (IsCandidate(child))
            {
                var holder = new Paragraph($"<w:p xmlns:w=\"{wordNamespace}\">{child.OuterXml}</w:p>");
                var typed = holder.FirstChild;

                // Still unknown as a paragraph's child, so nothing is gained by swapping it, and
                // leaving it keeps a second walk from handing out a different instance.
                if (typed is not (null or OpenXmlUnknownElement))
                {
                    typed.Remove();
                    element.ReplaceChild(typed, child);
                }
            }

            child = next;
        }
    }

    // A nested smart tag has no class either. It is retyped when the walk reaches it, and its
    // properties element is never a paragraph's child.
    static bool IsCandidate(OpenXmlElement child) =>
        child is OpenXmlUnknownElement {LocalName: not ("smartTag" or "smartTagPr")} &&
        child.NamespaceUri == wordNamespace;
}
