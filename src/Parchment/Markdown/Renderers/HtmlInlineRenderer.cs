class HtmlInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, HtmlInline>
{
    static readonly Dictionary<string, Action<RunProperties>> formatters = new(StringComparer.OrdinalIgnoreCase)
    {
        ["em"] = props => props.Append(new Italic()),
        ["i"] = props => props.Append(new Italic()),
        ["strong"] = props => props.Append(new Bold()),
        ["b"] = props => props.Append(new Bold()),
        ["u"] = props => props.Append(
            new Underline
            {
                Val = UnderlineValues.Single
            }),
        ["s"] = props => props.Append(new Strike()),
        ["del"] = props => props.Append(new Strike()),
        ["strike"] = props => props.Append(new Strike()),
        ["sub"] = props => props.Append(
            new VerticalTextAlignment
            {
                Val = VerticalPositionValues.Subscript
            }),
        ["sup"] = props => props.Append(
            new VerticalTextAlignment
            {
                Val = VerticalPositionValues.Superscript
            }),
    };

    protected override void Write(OpenXmlMarkdownRenderer renderer, HtmlInline inline)
    {
        var tag = inline.Tag;
        if (tag.Length == 0)
        {
            return;
        }

        if (TryParseTag(tag, out var name, out var isClosing, out var isSelfClosing, out var attributes))
        {
            if (string.Equals(name, "br", StringComparison.OrdinalIgnoreCase))
            {
                renderer.AddRun(new Run(new Break()));
                return;
            }

            // A character style is the one thing an otherwise unmapped tag can mean unambiguously in
            // Word, and <span class="..."> is the only way to ask for one mid-paragraph: a
            // {.StyleName} binds to emphasis, which imposes bold or italic alongside the style. It
            // also makes the inline path agree with the block one, where a span inside <h3> already
            // resolves through the html converter. Only the class is read - inline css stays that
            // converter's to interpret.
            if (string.Equals(name, "span", StringComparison.OrdinalIgnoreCase))
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
                    renderer.PushInlineHtmlFormat(name, StyleApplier(ReadClass(attributes)));
                }

                return;
            }

            if (formatters.TryGetValue(name, out var apply))
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

    /// <summary>
    /// Applies <paramref name="styleId"/> as a character style, or nothing when the span carried no
    /// class.
    /// </summary>
    /// <remarks>
    /// A classless span is still pushed rather than ignored: the closing tag pops by name, so
    /// skipping the push would leave <c>&lt;/span&gt;</c> popping whatever format encloses it.
    /// </remarks>
    static Action<RunProperties> StyleApplier(string? styleId)
    {
        if (styleId == null)
        {
            return _ =>
            {
            };
        }

        // Assigned through the typed property so rStyle leads the rPr sequence.
        return properties => properties.RunStyle = new()
        {
            Val = styleId
        };
    }

    /// <summary>
    /// The first class from a tag's attributes, or null when it carries none.
    /// </summary>
    /// <remarks>
    /// Only the first is read, for the same reason <see cref="MarkdownStyle"/> reads only the first
    /// of a <c>{.StyleName}</c>: a Word style is a single name, so the rest have nowhere to map.
    /// </remarks>
    static string? ReadClass(string attributes)
    {
        var remaining = attributes.AsSpan();
        while (true)
        {
            var index = remaining.IndexOf("class", StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                return null;
            }

            // Guards against matching the tail of another attribute name, which would read that
            // attribute's value as the style.
            var atBoundary = index == 0 || char.IsWhiteSpace(remaining[index - 1]);
            remaining = remaining[(index + 5)..];
            var value = remaining.TrimStart();
            if (atBoundary &&
                value.Length > 0 &&
                value[0] == '=')
            {
                return FirstClass(value[1..].TrimStart());
            }
        }
    }

    static string? FirstClass(ReadOnlySpan<char> value)
    {
        if (value.Length == 0)
        {
            return null;
        }

        var quote = value[0];
        if (quote is '"' or '\'')
        {
            value = value[1..];
            var close = value.IndexOf(quote);
            if (close < 0)
            {
                return null;
            }

            value = value[..close];
        }
        else
        {
            var end = 0;
            while (end < value.Length &&
                   !char.IsWhiteSpace(value[end]))
            {
                end++;
            }

            value = value[..end];
        }

        value = value.Trim();
        if (value.Length == 0)
        {
            return null;
        }

        var separator = value.IndexOfAny(' ', '\t');
        if (separator >= 0)
        {
            value = value[..separator];
        }

        return value.ToString();
    }

    static bool TryParseTag(string raw, out string name, out bool isClosing, out bool isSelfClosing, out string attributes)
    {
        name = "";
        attributes = "";
        isClosing = false;
        isSelfClosing = false;

        if (raw.Length < 3 ||
            raw[0] != '<' ||
            raw[^1] != '>')
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
        while (end < inner.Length &&
               (char.IsLetterOrDigit(inner[end]) || inner[end] == '-'))
        {
            end++;
        }

        if (end == 0)
        {
            return false;
        }

        if (end < inner.Length &&
            inner[end] != ' ' &&
            inner[end] != '\t')
        {
            return false;
        }

        name = inner[..end].ToString();
        attributes = inner[end..].Trim().ToString();
        return true;
    }
}
