/// <summary>
/// Writes an entity as the character it names.
/// </summary>
/// <remarks>
/// Markdig lifts an entity out of the literal around it into a node of its own, and writes nothing
/// for a leaf inline that no renderer claims — so with none registered <c>AT&amp;amp;T</c> reached
/// the document as "ATT", with no error to say so. The decoded text takes the literal's route from
/// there, which is what drops a character xml cannot carry (<c>&amp;#1;</c>) and makes a tab
/// (<c>&amp;#9;</c>) a real Word tab.
/// </remarks>
class HtmlEntityInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HtmlEntityInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, HtmlEntityInline inline)
    {
        var text = XmlChars.Strip(inline.Transcoded.AsSpan()).ToString();
        if (text.Length == 0)
        {
            return;
        }

        renderer.AddRun(LiteralInlineRenderer.BuildRun(text));
    }
}
