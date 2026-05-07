static class Filters
{
    public static void Register(FilterCollection filters)
    {
        filters.AddFilter("markdown", Markdown);
        filters.AddFilter("escape_xml", EscapeXml);
        filters.AddFilter("bullet_list", BulletList);
        filters.AddFilter("numbered_list", NumberedList);
    }

    static ValueTask<FluidValue> Markdown(FluidValue input, FilterArguments arguments, TemplateContext context)
    {
        var text = input.ToStringValue();
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

    static ValueTask<FluidValue> BulletList(FluidValue input, FilterArguments arguments, TemplateContext context) =>
        new(new ObjectValue(TokenValueHelpers.BulletList(Enumerate(input))));

    static ValueTask<FluidValue> NumberedList(FluidValue input, FilterArguments arguments, TemplateContext context) =>
        new(new ObjectValue(TokenValueHelpers.NumberedList(Enumerate(input))));

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
