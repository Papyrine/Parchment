static class DocxCloner
{
    public static MemoryStream ToWritableStream(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Copies <paramref name="source"/> into a fresh writable stream, reading it whole from the
    /// start whenever it can seek.
    /// </summary>
    /// <remarks>
    /// CopyTo covers MemoryStream too — its override writes straight out of the internal buffer, so
    /// one copy. Special-casing it as <c>ToWritableStream(source.ToArray())</c> cost a second copy
    /// of the whole package for no gain.
    ///
    /// The rewind keeps what that branch gave callers. ToArray ignored Position, so a template
    /// written into a MemoryStream and handed over without rewinding registered fine, while the
    /// same omission on a FileStream read from wherever it was left. Seeking first is the forgiving
    /// reading of both, and it is the one a caller who has just written the template expects.
    /// </remarks>
    public static MemoryStream ToWritableStream(Stream source)
    {
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        var stream = new MemoryStream();
        source.CopyTo(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Reads <paramref name="source"/> into an array, on the same terms as
    /// <see cref="ToWritableStream(Stream)"/> — whole, from the start, whenever it can seek.
    /// </summary>
    /// <remarks>
    /// The sibling of <see cref="ToWritableStream(Stream)"/>, and here beside it so the two agree on
    /// what a template stream means. They read the same sources and differ only in what they hand
    /// back, so a caller picks by what it needs to keep, not by which rules it gets.
    ///
    /// A MemoryStream can produce the array in one copy, which is why the special case earns its
    /// place here where <see cref="ToWritableStream(Stream)"/> could not justify it: there, the
    /// generic path already cost one copy and the branch added a second.
    /// </remarks>
    public static byte[] ToBytes(Stream source)
    {
        if (source.CanSeek)
        {
            source.Position = 0;
        }

        if (source is MemoryStream memory)
        {
            return memory.ToArray();
        }

        using var stream = new MemoryStream();
        source.CopyTo(stream);
        return stream.ToArray();
    }

    public static IEnumerable<(string uri, OpenXmlCompositeElement root)> EnumerateParts(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main == null)
        {
            yield break;
        }

        if (main.Document?.Body is { } body)
        {
            yield return (main.Uri.ToString(), body);
        }

        foreach (var headerPart in main.HeaderParts)
        {
            if (headerPart.Header is { } header)
            {
                yield return (headerPart.Uri.ToString(), header);
            }
        }

        foreach (var footerPart in main.FooterParts)
        {
            if (footerPart.Footer is { } footer)
            {
                yield return (footerPart.Uri.ToString(), footer);
            }
        }

        if (main.FootnotesPart?.Footnotes is { } footnotes)
        {
            yield return (main.FootnotesPart.Uri.ToString(), footnotes);
        }

        if (main.EndnotesPart?.Endnotes is { } endnotes)
        {
            yield return (main.EndnotesPart.Uri.ToString(), endnotes);
        }
    }
}
