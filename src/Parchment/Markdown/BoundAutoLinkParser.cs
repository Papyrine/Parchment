using Markdig.Extensions.AutoLinks;
using Markdig.Helpers;
using Markdig.Parsers;

/// <summary>
/// Markdig's bare-url autolink parser, made aware of the backslash escapes
/// <see cref="MarkdownEncoder"/> writes into a bound value.
/// </summary>
/// <remarks>
/// <para>
/// A template can write a url and bind its tail — <c>https://host/orders/{{ Id }}</c>. An autolink
/// is read verbatim, backslashes included, so the escapes that keep the value from being syntax
/// landed in the link's text and in its target, and the link was dead. They are undone here, once
/// the link has been delimited: by then the value can no longer be read as anything else.
/// </para>
/// <para>
/// Markdig also decides where the url ends without regard to escapes, so a value ending in
/// something a url may not end with (<c>.</c>, an unbalanced <c>)</c>, a <c>&lt;</c>) leaves the
/// link holding the backslash while the character it escaped falls outside, unescaped. The
/// backslash is handed back to the source, so that character is still escaped text after the link
/// rather than the start of emphasis, an entity or another autolink.
/// </para>
/// </remarks>
class BoundAutoLinkParser(AutoLinkOptions options) :
    AutoLinkParser(options)
{
    public override bool Match(InlineProcessor processor, ref StringSlice slice)
    {
        if (!base.Match(processor, ref slice))
        {
            return false;
        }

        if (processor.Inline is not LinkInline { Url: { } url, FirstChild: LiteralInline literal } link ||
            !url.Contains('\\'))
        {
            return true;
        }

        var text = literal.Content.AsSpan();
        if (MarkdownEncoder.EndsWithDanglingEscape(text))
        {
            text = text[..^1];
            url = url[..^1];
            link.Span.End--;
            slice.Start--;
        }

        link.Url = MarkdownEncoder.Unescape(url);
        literal.Content = new(MarkdownEncoder.Unescape(text));
        return true;
    }
}
