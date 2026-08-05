/// <summary>
/// Resolves a document's table of contents and page references against an
/// <see cref="IPageNumberResolver"/>, so the values are in the file before Word ever opens it.
/// </summary>
static class WordFieldResolution
{
    /// <summary>
    /// Give every table of contents its entries. Runs while the document is still being built,
    /// because entries added later would change the pagination that <see cref="Resolve"/> measures.
    /// </summary>
    public static void WriteTableOfContents(Body body)
    {
        foreach (var field in WordFieldRegion.Scan(body).Where(_ => _.IsTableOfContents))
        {
            WordTableOfContentsBuilder.WriteEntries(body, field);
        }
    }

    /// <summary>
    /// Measure the finished document and write the page numbers into it.
    /// </summary>
    /// <remarks>
    /// The resolver is handed a copy rather than the document itself: it is arbitrary caller code,
    /// and a renderer that seeks or disposes the stream it was given would otherwise corrupt the
    /// output it is measuring.
    /// </remarks>
    public static async Task Resolve(MemoryStream document, IPageNumberResolver resolver, Cancel cancel)
    {
        using var copy = new MemoryStream(document.ToArray(), false);
        var pages = await resolver.Resolve(copy, cancel);
        cancel.ThrowIfCancellationRequested();
        if (pages.Count == 0)
        {
            return;
        }

        document.Position = 0;
        using (var doc = WordprocessingDocument.Open(document, true))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null)
            {
                return;
            }

            foreach (var field in WordFieldRegion.Scan(body))
            {
                Apply(field, pages);
            }

            ApplyToEntries(body, pages);
            doc.Save();
        }

        document.Position = 0;
    }

    static void Apply(WordFieldRegion field, IReadOnlyDictionary<string, int> pages)
    {
        if (field.PageReferenceTarget is not { } target)
        {
            return;
        }

        if (pages.TryGetValue(target, out var page))
        {
            field.SetResult(page.ToString(CultureInfo.InvariantCulture));
        }
    }

    // A table-of-contents entry keeps its own page number in its last run, beside the hyperlink that
    // names which bookmark it points at. An entry the resolver could not place keeps its placeholder
    // and leaves the field marked for Word, so the reader still has a way to get a number.
    static void ApplyToEntries(Body body, IReadOnlyDictionary<string, int> pages)
    {
        foreach (var field in WordFieldRegion.Scan(body).Where(_ => _.IsTableOfContents))
        {
            var resolved = true;
            foreach (var hyperlink in EntriesOf(body, field))
            {
                var anchor = hyperlink.Anchor?.Value;
                var number = hyperlink.Elements<Run>().LastOrDefault()?.GetFirstChild<Text>();
                if (anchor == null || number == null)
                {
                    continue;
                }

                if (pages.TryGetValue(anchor, out var page))
                {
                    number.Text = page.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    resolved = false;
                }
            }

            if (resolved)
            {
                field.ClearDirty();
            }
        }
    }

    // The entries run from the paragraph the field starts in to the one its end marker closes.
    static IEnumerable<Hyperlink> EntriesOf(Body body, WordFieldRegion field)
    {
        var started = false;
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            if (paragraph == field.Paragraph)
            {
                started = true;
            }

            if (!started)
            {
                continue;
            }

            foreach (var hyperlink in paragraph.Elements<Hyperlink>())
            {
                yield return hyperlink;
            }

            if (paragraph.Elements<Run>().Any(_ => _.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End))
            {
                yield break;
            }
        }
    }
}
