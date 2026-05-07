class HtmlInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HtmlInline>
{
    static readonly Dictionary<string, Action<RunProperties>> Formatters = new(StringComparer.OrdinalIgnoreCase)
    {
        ["em"] = props => props.Append(new Italic()),
        ["i"] = props => props.Append(new Italic()),
        ["strong"] = props => props.Append(new Bold()),
        ["b"] = props => props.Append(new Bold()),
        ["u"] = props => props.Append(new Underline { Val = UnderlineValues.Single }),
        ["s"] = props => props.Append(new Strike()),
        ["del"] = props => props.Append(new Strike()),
        ["strike"] = props => props.Append(new Strike()),
        ["sub"] = props => props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }),
        ["sup"] = props => props.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
    };

    protected override void Write(OpenXmlMarkdownRenderer renderer, HtmlInline inline)
    {
        var tag = inline.Tag;
        if (tag.Length == 0)
        {
            return;
        }

        if (TryParseTag(tag, out var name, out var isClosing, out var isSelfClosing))
        {
            if (string.Equals(name, "br", StringComparison.OrdinalIgnoreCase))
            {
                renderer.AddRun(new Run(new Break()));
                return;
            }

            if (Formatters.TryGetValue(name, out var apply))
            {
                if (isSelfClosing)
                {
                    return;
                }

                if (isClosing)
                {
                    renderer.PopInlineHtmlFormat(name);
                }
                else
                {
                    renderer.PushInlineHtmlFormat(name, apply);
                }

                return;
            }
        }

        renderer.AddRun(
            new Run(
                new Text(tag)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
    }

    static bool TryParseTag(string raw, out string name, out bool isClosing, out bool isSelfClosing)
    {
        name = "";
        isClosing = false;
        isSelfClosing = false;

        if (raw.Length < 3 || raw[0] != '<' || raw[^1] != '>')
        {
            return false;
        }

        var inner = raw.AsSpan(1, raw.Length - 2);

        if (inner.Length > 0 && inner[0] == '/')
        {
            isClosing = true;
            inner = inner[1..];
        }

        if (inner.Length > 0 && inner[^1] == '/')
        {
            isSelfClosing = true;
            inner = inner[..^1].TrimEnd();
        }

        var end = 0;
        while (end < inner.Length && (char.IsLetterOrDigit(inner[end]) || inner[end] == '-'))
        {
            end++;
        }

        if (end == 0)
        {
            return false;
        }

        if (end < inner.Length && inner[end] != ' ' && inner[end] != '\t')
        {
            return false;
        }

        name = inner[..end].ToString();
        return true;
    }
}
