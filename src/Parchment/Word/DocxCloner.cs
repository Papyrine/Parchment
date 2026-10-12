static class DocxCloner
{
    public static MemoryStream ToWritableStream(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }

    public static MemoryStream ToWritableStream(Stream source)
    {
        if (source is MemoryStream ms)
        {
            return ToWritableStream(ms.ToArray());
        }

        var stream = new MemoryStream();
        source.CopyTo(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// The parts that can carry tokens, each with its root element.
    /// </summary>
    /// <remarks>
    /// Reading a part's root makes the SDK parse the part, and a parsed part is written out again
    /// on save whether or not anything in it changed. <paramref name="only"/> names, by uri, the
    /// parts a caller has work in, so the rest stay as the bytes the package was opened with.
    /// </remarks>
    public static IEnumerable<(string uri, OpenXmlCompositeElement root)> EnumerateParts(WordprocessingDocument doc, IReadOnlySet<string>? only = null)
    {
        var main = doc.MainDocumentPart;
        if (main == null)
        {
            yield break;
        }

        var mainUri = main.Uri.ToString();
        if (Wanted(only, mainUri) &&
            main.Document?.Body is { } body)
        {
            yield return (mainUri, body);
        }

        foreach (var headerPart in main.HeaderParts)
        {
            var uri = headerPart.Uri.ToString();
            if (Wanted(only, uri) &&
                headerPart.Header is { } header)
            {
                yield return (uri, header);
            }
        }

        foreach (var footerPart in main.FooterParts)
        {
            var uri = footerPart.Uri.ToString();
            if (Wanted(only, uri) &&
                footerPart.Footer is { } footer)
            {
                yield return (uri, footer);
            }
        }

        if (main.FootnotesPart is { } footnotesPart)
        {
            var uri = footnotesPart.Uri.ToString();
            if (Wanted(only, uri) &&
                footnotesPart.Footnotes is { } footnotes)
            {
                yield return (uri, footnotes);
            }
        }

        if (main.EndnotesPart is { } endnotesPart)
        {
            var uri = endnotesPart.Uri.ToString();
            if (Wanted(only, uri) &&
                endnotesPart.Endnotes is { } endnotes)
            {
                yield return (uri, endnotes);
            }
        }
    }

    static bool Wanted(IReadOnlySet<string>? only, string uri) =>
        only == null ||
        only.Contains(uri);
}
