namespace Parchment;

/// <summary>
/// Turns the newlines a substituted value carried into real <c>w:br</c> elements, so a line break
/// in bound content is a line break in the document.
/// </summary>
/// <remarks>
/// <para>
/// A newline written into a <c>w:t</c> is not a line break — Word lays it out as whitespace, so a
/// multi-line value silently arrived as one run of text. Only a <c>w:br</c> beside it breaks the
/// line.
/// </para>
/// <para>
/// Substitution is text-based and offset-driven: <c>ParagraphText</c> maps character offsets onto
/// the <c>w:t</c> elements that own them, and <c>ScopeTreeRunner</c> applies a paragraph's tokens in
/// reverse-offset order so each replacement leaves the offsets of the ones still to come intact.
/// Splitting a <c>w:t</c> at substitution time would invalidate that map mid-walk, so the newline is
/// parked as <see cref="Marker"/> — which costs one character, like the newline it replaces, and so
/// disturbs nothing — and the split happens once per part after every substitution has run.
/// </para>
/// <para>
/// The markdown flow's body does not come through here. It has no <c>w:t</c> to split at
/// substitution time, so <see cref="MarkdownEncoder"/> writes an inline <c>&lt;br /&gt;</c> into the
/// markdown instead and <c>HtmlInlineRenderer</c> renders it to the same <c>w:br</c>. A markdown
/// template's headers and footers are docx parts, so those do come through here.
/// </para>
/// </remarks>
static class LineBreaks
{
    /// <summary>
    /// Stands in for a line break between substitution and <see cref="Apply"/>.
    /// </summary>
    /// <remarks>
    /// A C0 control that <see cref="OpenXmlKit.XmlChars"/> strips, and it is written only after that
    /// strip has run. So a marker in the tree can only be one Parchment put there — never a
    /// character that arrived in a model value — which is what makes <see cref="Apply"/> safe to
    /// run across a whole part rather than having to track which text came from where.
    /// </remarks>
    public const char Marker = '\u0001';

    /// <summary>
    /// Swaps the newlines in a substituted value for <see cref="Marker"/>. A CRLF is one marker.
    /// </summary>
    public static CharSpan Mark(CharSpan value)
    {
        var first = value.IndexOfAny('\r', '\n');
        if (first < 0)
        {
            return value;
        }

        var builder = new StringBuilder(value.Length);
        builder.Append(value[..first]);

        for (var index = first; index < value.Length; index++)
        {
            var character = value[index];
            if (character is not ('\r' or '\n'))
            {
                builder.Append(character);
                continue;
            }

            if (character == '\r' &&
                index + 1 < value.Length &&
                value[index + 1] == '\n')
            {
                index++;
            }

            builder.Append(Marker);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Replaces every marker under <paramref name="root"/> with a <c>w:br</c>.
    /// </summary>
    public static void Apply(OpenXmlCompositeElement root)
    {
        // Collected before splitting: the split inserts siblings, and mutating the tree while
        // enumerating Descendants would revisit what it just created.
        List<Text>? marked = null;
        foreach (var text in root.Descendants<Text>())
        {
            if (text.Text.Contains(Marker))
            {
                marked ??= [];
                marked.Add(text);
            }
        }

        if (marked == null)
        {
            return;
        }

        foreach (var text in marked)
        {
            Split(text);
        }
    }

    static void Split(Text text)
    {
        var segments = text.Text.Split(Marker);

        // The first segment stays in the existing element, so its run properties and its position
        // among the paragraph's other runs are untouched. The breaks and the segments after them
        // go into the same run for the same reason — a w:br is a sibling of w:t, so nothing has to
        // be cloned to keep the formatting.
        text.Text = segments[0];
        text.Space = SpaceProcessingModeValues.Preserve;

        var parent = text.Parent;
        if (parent == null)
        {
            return;
        }

        OpenXmlElement previous = text;
        for (var index = 1; index < segments.Length; index++)
        {
            previous = parent.InsertAfter(new Break(), previous);

            var segment = segments[index];
            if (segment.Length == 0)
            {
                continue;
            }

            previous = parent.InsertAfter(
                new Text(segment)
                {
                    Space = SpaceProcessingModeValues.Preserve
                },
                previous);
        }
    }
}
