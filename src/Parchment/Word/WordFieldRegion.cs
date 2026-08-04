/// <summary>
/// One Word field located in a document: its instruction, and the runs that hold the cached result
/// Word displays until it recalculates.
/// </summary>
/// <remarks>
/// A field is not an element but a run sequence — <c>begin</c>, the instruction, <c>separate</c>,
/// the result, <c>end</c> — so finding one means walking runs and pairing the markers. Nesting is
/// possible in Word and does not arise here: Parchment emits no field inside another.
/// </remarks>
class WordFieldRegion
{
    public required Paragraph Paragraph { get; init; }
    public required Run Begin { get; init; }
    public required string Instruction { get; init; }
    public required IReadOnlyList<Run> Result { get; init; }

    /// <summary>The bookmark a PAGEREF names, or null for any other field.</summary>
    public string? PageReferenceTarget =>
        Instruction.TrimStart().StartsWith("PAGEREF ", StringComparison.Ordinal)
            ? Instruction.Trim().Split(' ')[1]
            : null;

    public bool IsTableOfContents =>
        Instruction.TrimStart().StartsWith("TOC", StringComparison.Ordinal);

    /// <summary>
    /// Replace what the field displays, and stop asking Word to recalculate it — the value is now
    /// as good as the resolver could make it, and a dirty field costs the reader a prompt on open.
    /// </summary>
    public void SetResult(string text)
    {
        foreach (var run in Result.Skip(1))
        {
            run.Remove();
        }

        if (Result.Count == 0)
        {
            return;
        }

        var first = Result[0];

        first.RemoveAllChildren<Text>();
        first.Append(
            new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        ClearDirty();
    }

    public void ClearDirty()
    {
        var fieldChar = Begin.GetFirstChild<FieldChar>();
        if (fieldChar?.Dirty != null)
        {
            fieldChar.Dirty = null;
        }
    }

    /// <summary>
    /// Every field under <paramref name="root"/>, in document order.
    /// </summary>
    /// <remarks>
    /// Walked over runs rather than per paragraph, because a field is not confined to one: a built
    /// table of contents opens in the paragraph holding its first entry and closes in the paragraph
    /// holding its last.
    /// </remarks>
    public static List<WordFieldRegion> Scan(OpenXmlElement root)
    {
        Run? begin = null;
        Paragraph? paragraph = null;
        var instruction = new StringBuilder();
        List<Run>? result = null;
        var found = new List<WordFieldRegion>();

        foreach (var run in root.Descendants<Run>())
        {
            var marker = run.GetFirstChild<FieldChar>()?.FieldCharType?.Value;
            if (marker == FieldCharValues.Begin)
            {
                begin = run;
                paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
                instruction.Clear();
                result = null;
                continue;
            }

            if (begin == null ||
                paragraph == null)
            {
                continue;
            }

            if (marker == FieldCharValues.Separate)
            {
                result = [];
                continue;
            }

            if (marker == FieldCharValues.End)
            {
                found.Add(new()
                {
                    Paragraph = paragraph,
                    Begin = begin,
                    Instruction = instruction.ToString().Trim(),
                    Result = result ?? []
                });
                begin = null;
                continue;
            }

            if (result != null)
            {
                result.Add(run);
                continue;
            }

            foreach (var code in run.Elements<FieldCode>())
            {
                instruction.Append(code.Text);
            }
        }

        return found;
    }
}
