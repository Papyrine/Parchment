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

    static void WriteLink(OpenXmlMarkdownRenderer renderer, LinkInline inline)
    {
        var url = inline.Url ?? string.Empty;
        if (string.IsNullOrEmpty(url))
        {
            renderer.WriteChildren(inline);
            return;
        }

        var relId = renderer.MainPart.AddHyperlinkRelationship(new(url, UriKind.RelativeOrAbsolute), true).Id;

        var top = renderer.Top;
        var before = top.CurrentRuns.Count;
        renderer.WriteChildren(inline);
        var produced = top.CurrentRuns.Skip(before).ToList();

        var hyperlink = new Hyperlink
        {
            Id = relId
        };
        var styleId = MarkdownStyle.Resolve(inline) ?? "Hyperlink";
        foreach (var run in produced)
        {
            if (run is Run runElement)
            {
                runElement.RunProperties ??= new();
                runElement.RunProperties.Append(
                    new RunStyle
                    {
                        Val = styleId
                    });
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
