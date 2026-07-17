static class Filters
{
    public static void Register(FilterCollection filters)
    {
        filters.AddFilter("markdown", Markdown);
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
            // through.
            return new(new Fluid.Values.StringValue(text));
        }

        return new(new ObjectValue(new MarkdownToken(text)));
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

        return new(new Fluid.Values.StringValue(builder.ToString()));
    }

    static ValueTask<FluidValue> BulletList(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        if (IsMarkdownFlow(context))
        {
            return new(new Fluid.Values.StringValue(BuildList(input, static (_, _) => "- ")));
        }

        return new(new ObjectValue(TokenValueHelpers.BulletList(Enumerate(input))));
    }

    static ValueTask<FluidValue> NumberedList(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        if (IsMarkdownFlow(context))
        {
            return new(new Fluid.Values.StringValue(BuildList(input, static (_, index) => $"{index + 1}. ")));
        }

        return new(new ObjectValue(TokenValueHelpers.NumberedList(Enumerate(input))));
    }

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
            builder.Append(item);
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
