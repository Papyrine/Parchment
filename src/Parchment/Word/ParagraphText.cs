/// <summary>
/// Builds the flat InnerText of a paragraph plus a map from character offset → (Text element, char index within).
/// Used to locate tokens that straddle run boundaries and apply substitutions back to the XML tree.
/// </summary>
class ParagraphText
{
    List<TextSpan> spans;
    string innerText;

    ParagraphText(List<TextSpan> spans, string innerText)
    {
        this.spans = spans;
        this.innerText = innerText;
    }

    public string InnerText => innerText;

    public static ParagraphText Build(Paragraph paragraph)
    {
        // Fast path: most paragraphs (especially loop-body paragraphs) have ≤1 Text descendant.
        // Track the first one in locals; only allocate StringBuilder when a second appears.
        // Skips both the StringBuilder allocation and the trailing ToString() in the common case,
        // and reuses text.Text directly as innerText.
        var spans = new List<TextSpan>(4);
        StringBuilder? builder = null;
        string? singleValue = null;

        foreach (var text in paragraph.Descendants<Text>())
        {
            var value = text.Text;
            var offset = builder?.Length ?? singleValue?.Length ?? 0;
            spans.Add(new(offset, value.Length, text));

            if (builder != null)
            {
                builder.Append(value);
            }
            else if (singleValue == null)
            {
                singleValue = value;
            }
            else
            {
                builder = new(singleValue.Length + value.Length);
                builder.Append(singleValue);
                builder.Append(value);
                singleValue = null;
            }
        }

        var innerText = builder?.ToString() ?? singleValue ?? string.Empty;
        return new(spans, innerText);
    }

    /// <summary>
    /// Replaces a character range in the InnerText with the given replacement, preserving the formatting
    /// of the run that owns the first character of the range. Intermediate runs inside the range are deleted.
    /// </summary>
    public void Replace(int offset, int length, string replacement)
    {
        // Marked after the strip, never before: the marker is one of the characters the strip
        // removes, which is what guarantees a marker in the tree came from here and not from a
        // model value. See LineBreaks.
        var cleaned = LineBreaks.Mark(XmlChars.Strip(replacement));

        if (length < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(length));
        }

        if (offset < 0 || offset > innerText.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(offset));
        }

        if (length == 0)
        {
            return;
        }

        // Fast path: single-span paragraph. Skips two FindSpan scans and the cross-span
        // branching. The token offset/length are guaranteed to fall within the only span.
        if (spans.Count == 1)
        {
            var only = spans[0];
            var source = only.Text.Text;
            only.Text.Text = string.Concat(source.AsSpan(0, offset), cleaned, source.AsSpan(offset + length));
            only.Text.Space = SpaceProcessingModeValues.Preserve;
            return;
        }

        var end = offset + length;
        var first = FindSpan(offset, preferEnd: false);
        var last = FindSpan(end - 1, preferEnd: true);

        if (first.index == last.index)
        {
            ReplaceWithin(first, offset, length, cleaned);
            return;
        }

        // Different spans — put all replacement text into the first span, and trim the rest.
        var firstSpan = spans[first.index];
        var localStart = offset - firstSpan.Offset;
        var firstText = firstSpan.Text.Text;
        var newFirstText = string.Concat(firstText.AsSpan(0, localStart), cleaned);
        firstSpan.Text.Text = newFirstText;
        firstSpan.Text.Space = SpaceProcessingModeValues.Preserve;

        var lastSpan = spans[last.index];
        var localEnd = end - lastSpan.Offset;
        var lastText = lastSpan.Text.Text;
        lastSpan.Text.Text = lastText[localEnd..];
        if (lastSpan.Text.Text.Length > 0)
        {
            lastSpan.Text.Space = SpaceProcessingModeValues.Preserve;
        }

        // Remove any intermediate text elements entirely (empty them so downstream XML stays valid).
        for (var i = first.index + 1; i < last.index; i++)
        {
            spans[i].Text.Text = string.Empty;
        }
    }

    void ReplaceWithin(SpanRef reference, int offset, int length, CharSpan replacement)
    {
        var span = spans[reference.index];
        var local = offset - span.Offset;
        var source = span.Text.Text;
        var updated = string.Concat(source.AsSpan(0, local), replacement, source.AsSpan(local + length));
        span.Text.Text = updated;
        span.Text.Space = SpaceProcessingModeValues.Preserve;
    }

    SpanRef FindSpan(int absoluteOffset, bool preferEnd)
    {
        for (var i = 0; i < spans.Count; i++)
        {
            var span = spans[i];
            var end = span.Offset + span.Length;
            if (absoluteOffset >= span.Offset && absoluteOffset < end)
            {
                return new(i);
            }

            if (preferEnd && absoluteOffset == end && i == spans.Count - 1)
            {
                return new(i);
            }
        }

        throw new InvalidOperationException($"Offset {absoluteOffset} is not contained in any span (innerText length {innerText.Length}).");
    }

    readonly record struct SpanRef(int index);
}

readonly record struct TextSpan(int Offset, int Length, Text Text);
