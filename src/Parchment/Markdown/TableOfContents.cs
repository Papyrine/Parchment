/// <summary>
/// Turns a paragraph whose whole content is <c>[TOC]</c> into a Word table-of-contents field.
/// </summary>
/// <remarks>
/// <c>[TOC]</c> is the spelling several markdown dialects already use, and markdown itself does
/// nothing with it: with no <c>(url)</c> or <c>[label]</c> following, it parses as the literal text
/// "[TOC]". So the marker costs no syntax and a template that renders to html elsewhere degrades to
/// showing the marker rather than breaking.
/// <para>
/// The depth comes from a generic-attribute property — <c>[TOC]{levels=1}</c> lists Heading1 only.
/// Word's own default is three levels, so that is the default here.
/// </para>
/// </remarks>
static class TableOfContents
{
    const string Marker = "[TOC]";
    const int DefaultLevels = 3;
    const int MaxLevels = 9;

    // Shown until Word recalculates the field. Word updates a dirty TOC on open, so this is what a
    // reader that does not support fields displays.
    const string Placeholder = "Right-click and choose \"Update Field\" to build the table of contents.";

    public static bool TryWrite(OpenXmlMarkdownRenderer renderer, ParagraphBlock block, ParagraphProperties? properties)
    {
        if (!IsMarker(block))
        {
            return false;
        }

        foreach (var run in WordFields.TableOfContents(ResolveLevels(block), Placeholder))
        {
            renderer.AddRun(run);
        }

        renderer.FlushParagraph(properties);
        return true;
    }

    // An unresolved "[TOC]" reaches the renderer as plain text, but not necessarily as one run of
    // it — the link parser splits the brackets off from the word between them. So the test is that
    // the paragraph holds text and nothing else, and that the text is the marker.
    static bool IsMarker(ParagraphBlock block)
    {
        var child = block.Inline?.FirstChild;
        if (child == null)
        {
            return false;
        }

        var builder = new StringBuilder();
        while (child != null)
        {
            if (child is not LiteralInline literal)
            {
                return false;
            }

            builder.Append(literal.Content.AsSpan());
            child = child.NextSibling;
        }

        return builder.ToString().Trim() == Marker;
    }

    static int ResolveLevels(ParagraphBlock block)
    {
        var properties = block.TryGetAttributes()?.Properties;
        if (properties == null)
        {
            return DefaultLevels;
        }

        foreach (var (key, value) in properties)
        {
            if (!string.Equals(key, "levels", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var levels))
            {
                return Math.Clamp(levels, 1, MaxLevels);
            }
        }

        return DefaultLevels;
    }
}
