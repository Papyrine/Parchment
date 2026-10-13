using Markdig.Extensions.SmartyPants;

class LinkInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, LinkInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, LinkInline inline)
    {
        if (inline.IsImage)
        {
            WriteImage(renderer, inline);
            return;
        }

        WriteLink(renderer, inline);
    }

    // Shown until Word recalculates a PAGEREF. Any digit would do; "1" is what Word itself caches
    // for a reference it has not resolved yet.
    const string PageReferencePlaceholder = "1";

    static void WriteLink(OpenXmlMarkdownRenderer renderer, LinkInline inline)
    {
        var url = inline.Url ?? string.Empty;
        if (string.IsNullOrEmpty(url))
        {
            renderer.WriteChildren(inline);
            return;
        }

        // A "#name" url addresses somewhere in this document, so it is an anchor on the hyperlink
        // rather than a relationship to an external target. With no link text there is nothing to
        // click either, and the only useful thing left to say about a place in the document is which
        // page it is on — so an empty internal link becomes a PAGEREF instead.
        if (url[0] == '#')
        {
            var bookmark = renderer.BookmarkNames.For(url[1..]);
            if (inline.FirstChild == null)
            {
                foreach (var run in WordFields.PageReference(bookmark, PageReferencePlaceholder))
                {
                    renderer.AddRun(run);
                }

                return;
            }

            WriteHyperlink(renderer, inline, new()
            {
                Anchor = bookmark
            });
            return;
        }

        if (!RelationshipIds.TryHyperlink(renderer.Part, url, out var relId))
        {
            WriteUnlinked(renderer, inline, url);
            return;
        }

        WriteHyperlink(renderer, inline, new()
        {
            Id = relId
        });
    }

    // A destination no relationship can point at — one that is not a url at all, or a relative one,
    // which has nothing to be relative to here — leaves the link as its text followed by the
    // destination in parentheses, so the reader still has both. This is what OpenXmlHtml writes for
    // an <a href> it cannot resolve. Text that already is the destination is not repeated.
    static void WriteUnlinked(OpenXmlMarkdownRenderer renderer, LinkInline inline, string url)
    {
        url = XmlChars.Strip(url.AsSpan()).ToString();
        if (inline.FirstChild == null)
        {
            renderer.AddRun(LiteralInlineRenderer.BuildRun(url));
            return;
        }

        renderer.WriteChildren(inline);
        if (!ExtractAlt(inline).AsSpan().Trim().Equals(url, StringComparison.Ordinal))
        {
            renderer.AddRun(LiteralInlineRenderer.BuildRun($" ({url})"));
        }
    }

    static void WriteHyperlink(OpenXmlMarkdownRenderer renderer, LinkInline inline, Hyperlink hyperlink)
    {
        var top = renderer.Top;
        var before = top.CurrentRuns.Count;
        renderer.WriteChildren(inline);
        var produced = top.CurrentRuns.Skip(before).ToList();

        var styleId = MarkdownStyle.Resolve(inline) ?? "Hyperlink";
        foreach (var run in produced)
        {
            if (run is Run runElement)
            {
                // Where the children have already given the run properties, the style is assigned
                // through the typed property, which puts rStyle first in the rPr sequence. A run
                // carries one character style: one a child asked for by name ("`code`{.Mono}", a
                // classed span) is the more specific and is kept. Plain link text has no properties
                // yet, and they are built holding the style.
                if (runElement.GetFirstChild<RunProperties>() is { } properties)
                {
                    properties.RunStyle ??= new()
                    {
                        Val = styleId
                    };
                }
                else
                {
                    runElement.PrependChild(
                        new RunProperties(
                            new RunStyle
                            {
                                Val = styleId
                            }));
                }

                hyperlink.Append(runElement);
            }
            else
            {
                hyperlink.Append(run);
            }
        }

        top.CurrentRuns.RemoveRange(before, top.CurrentRuns.Count - before);
        renderer.AddRun(hyperlink);
    }

    // Synthesize an <img> tag and delegate to OpenXmlHtml so drawing/blip plumbing — including
    // data: URIs, file paths, and http(s) sources — is resolved by ImageResolver in one place.
    static void WriteImage(OpenXmlMarkdownRenderer renderer, LinkInline inline)
    {
        var url = inline.Url;
        if (string.IsNullOrWhiteSpace(url))
        {
            renderer.WriteChildren(inline);
            return;
        }

        var alt = ExtractAlt(inline);
        var html = $"<img src=\"{HtmlEscape(url)}\" alt=\"{HtmlEscape(alt)}\" />";

        var settings = renderer.ImagePolicies.BuildSettings();
        var elements = WordHtmlConverter.ToElements(html, renderer.MainPart, settings);
        renderer.MayHoldBookmarks = true;
        foreach (var element in elements)
        {
            if (element is Paragraph paragraph)
            {
                foreach (var run in paragraph.ChildElements.OfType<Run>().ToList())
                {
                    run.Remove();
                    renderer.AddRun(run);
                }
            }
        }
    }

    static string ExtractAlt(LinkInline inline)
    {
        var builder = new StringBuilder();
        Walk(inline, builder);
        return builder.ToString();
    }

    // Alt text is plain, so each inline gives the text it would have rendered as and none of its
    // formatting. It is what shows when the image cannot be loaded, so an inline left out here is
    // text missing from the document the same way one with no renderer is.
    static void Walk(ContainerInline container, StringBuilder builder)
    {
        var child = container.FirstChild;
        while (child != null)
        {
            switch (child)
            {
                case LiteralInline literal:
                    builder.Append(literal.Content);
                    break;
                case HtmlEntityInline entity:
                    builder.Append(XmlChars.Strip(entity.Transcoded.AsSpan()));
                    break;
                case SmartyPant pant:
                    builder.Append(SmartyPantInlineRenderer.Glyph(pant.Type));
                    break;
                case CodeInline code:
                    builder.Append(code.ContentSpan);
                    break;
                case AutolinkInline autolink:
                    builder.Append(autolink.Url);
                    break;
                case LineBreakInline:
                    builder.Append(' ');
                    break;
                case DelimiterInline delimiter:
                    builder.Append(delimiter.ToLiteral());
                    Walk(delimiter, builder);
                    break;
                case ContainerInline inner:
                    Walk(inner, builder);
                    break;
            }
            child = child.NextSibling;
        }
    }

    static string HtmlEscape(string value) =>
        value
            .Replace("&", "&amp;")
            .Replace("\"", "&quot;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;");
}
