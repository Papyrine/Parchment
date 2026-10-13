/// <summary>
/// Writes a delimiter that never found its match as the text it was typed as, then what it holds.
/// </summary>
/// <remarks>
/// <para>
/// A delimiter is a container: everything after an open <c>[</c> becomes its children while the
/// parser waits for the <c>]</c>. One that closes is replaced — by a link, or by a literal when it
/// turns out not to be one — but one that never closes stays in the tree as it is. Markdig writes
/// the children of a container no renderer claims, so what followed the bracket rendered and only
/// the bracket was lost: <c>see [note</c> became "see note".
/// </para>
/// <para>
/// Registered for the base type, so the same holds for whichever delimiter a parser leaves behind
/// and not only the link parser's.
/// </para>
/// </remarks>
class DelimiterInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, DelimiterInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, DelimiterInline inline)
    {
        var literal = inline.ToLiteral();
        if (literal.Length > 0)
        {
            renderer.AddRun(LiteralInlineRenderer.BuildRun(literal));
        }

        renderer.WriteChildren(inline);
    }
}
