/// <summary>
/// Writes a character reference (<c>&amp;amp;</c>, <c>&amp;#32;</c>) as the character it names.
/// </summary>
/// <remarks>
/// Markdig keeps a reference as its own inline rather than folding it into the literal around it,
/// and an inline with no renderer is skipped — so without this the character goes missing from the
/// document. <see cref="MarkdownEncoder"/> depends on it: a reference is how the leading space or
/// tab of a bound value is carried past the block parser.
/// </remarks>
class HtmlEntityInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HtmlEntityInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, HtmlEntityInline inline)
    {
        var content = inline.Transcoded.AsSpan();
        if (content.Length == 0)
        {
            return;
        }

        var text = XmlChars.Strip(content).ToString();
        renderer.AddRun(LiteralInlineRenderer.BuildRun(text));
    }
}
