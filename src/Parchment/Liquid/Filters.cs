static class Filters
{
    public static void Register(FilterCollection filters)
    {
        filters.AddFilter("markdown", Markdown);
        filters.AddFilter("html", Html);
        filters.AddFilter("escape_xml", EscapeXml);
        filters.AddFilter("bullet_list", BulletList);
        filters.AddFilter("numbered_list", NumberedList);
    }

    // A filter returns a FluidValue, so it bypasses the TokenValue value converter on
    // SharedFluid.MarkdownOptions (that only runs when Fluid creates a value from a CLR object).
    // The markdown flow therefore has to be detected here and the markdown source emitted directly;
    // otherwise the token reaches the writer and Fluid writes its type name into the document.
    static bool IsMarkdownFlow(TemplateContext context) =>
        ReferenceEquals(context.Options, SharedFluid.MarkdownOptions);

    static ValueTask<FluidValue> Markdown(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        var text = input.ToStringValue();
        if (IsMarkdownFlow(context))
        {
            // The result is about to be parsed as markdown anyway, so the source passes straight
            // through — unencoded, since asking for the markdown filter is asking for the value to
            // be read as syntax. This is the explicit form of what MarkdownEncoder turns off by
            // default for every other substitution.
            return new(new Fluid.Values.StringValue(text, encode: false));
        }

        return new(new ObjectValue(new MarkdownToken(text)));
    }

    // The counterpart to `markdown`, and unlike it not a pass-through: markdown source written into
    // a markdown template is already what it will be parsed as, whereas html written there is only
    // html if Markdig happens to agree. So the value is parked and converted after the parse — see
    // MarkdownTokenBlocks — which is what makes `| html` mean the same thing in both flows.
    static ValueTask<FluidValue> Html(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        var text = input.ToStringValue();
        if (IsMarkdownFlow(context))
        {
            return new(new Fluid.Values.StringValue(MarkdownTokenBlocks.Register(text), encode: false));
        }

        return new(new ObjectValue(new HtmlToken(text)));
    }

    static ValueTask<FluidValue> EscapeXml(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        var text = input.ToStringValue();
        var builder = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            switch (c)
            {
                case '<':
                    builder.Append("&lt;");
                    break;
                case '>':
                    builder.Append("&gt;");
                    break;
                case '&':
                    builder.Append("&amp;");
                    break;
                case '"':
                    builder.Append("&quot;");
                    break;
                case '\'':
                    builder.Append("&apos;");
                    break;
                default:
                    builder.Append(c);
                    break;
            }
        }

        // Unencoded: escaping the entities back into literal text would undo the filter. Asking for
        // xml escaping says the value is headed for a markup context the template is assembling, so
        // MarkdownEncoder stays out of it.
        return new(new Fluid.Values.StringValue(builder.ToString(), encode: false));
    }

    static ValueTask<FluidValue> BulletList(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        if (IsMarkdownFlow(context))
        {
            return new(new Fluid.Values.StringValue(BuildList(input, static (_, _) => "- "), encode: false));
        }

        return new(new ObjectValue(TokenValueHelpers.BulletList(Enumerate(input))));
    }

    static ValueTask<FluidValue> NumberedList(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        if (IsMarkdownFlow(context))
        {
            return new(new Fluid.Values.StringValue(BuildList(input, static (_, index) => $"{index + 1}. "), encode: false));
        }

        return new(new ObjectValue(TokenValueHelpers.NumberedList(Enumerate(input))));
    }

    // Markdown flow only — the docx flow builds real list paragraphs via TokenValueHelpers. The
    // result is returned unencoded because the markers and the newlines between them have to stay
    // syntax, so each item is escaped here instead: an item reading "1. First" or carrying a "|"
    // is text, not a nested list or a table cell.
    static string BuildList(FluidValue input, Func<string, int, string> marker)
    {
        var builder = new StringBuilder();
        var index = 0;
        foreach (var item in Enumerate(input))
        {
            if (index > 0)
            {
                builder.Append('\n');
            }

            builder.Append(marker(item, index));
            builder.Append(MarkdownEncoder.EscapeValue(item));
            index++;
        }

        return builder.ToString();
    }

    static IEnumerable<string> Enumerate(FluidValue input)
    {
        if (input is ArrayValue array)
        {
            foreach (var item in array.Values)
            {
                yield return item.ToStringValue();
            }

            yield break;
        }

        var value = input.ToObjectValue();

        if (value is IEnumerable<object?> objects)
        {
            foreach (var item in objects)
            {
                yield return item?.ToString() ?? string.Empty;
            }

            yield break;
        }

        if (value is IEnumerable raw and not string)
        {
            foreach (var item in raw)
            {
                yield return item?.ToString() ?? string.Empty;
            }

            yield break;
        }

        yield return input.ToStringValue();
    }
}
