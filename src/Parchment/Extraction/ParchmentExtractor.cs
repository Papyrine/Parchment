namespace Parchment;

/// <summary>
/// Reads <c>[EditableField]</c> values back out of a docx produced by a Parchment template —
/// the other half of two-way binding. Content controls are matched by their <c>w:tag</c>
/// (the dotted model path), so no <see cref="TemplateStore"/> or template registration is
/// required. Checkbox / date / dropdown controls are read from canonical control state; string
/// and numeric fields read display text, with numerics parsed using the <c>culture</c> argument —
/// which must match the render culture (Parchment renders with Fluid's default invariant
/// culture unless customized, so the default here is also <see cref="CultureInfo.InvariantCulture"/>).
/// </summary>
public static class ParchmentExtractor
{
    // Extract is called per-document (often per-request); cache the reflection-built map per
    // model type. The SG-precompiled cache inside EditableMap.Build makes this a double layer on
    // the SG path, but it also covers the reflection fallback path, which Build alone does not.
    static ConcurrentDictionary<Type, EditableMap> maps = new();

    public static ExtractResult<TModel> Extract<TModel>(string path, CultureInfo? culture = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var file = File.OpenRead(path);
        return Extract<TModel>(file, culture);
    }

    public static ExtractResult<TModel> Extract<TModel>(Stream docx, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(docx);

        var map = maps.GetOrAdd(typeof(TModel), static type => EditableMap.Build(type, type.Name));
        if (map.IsEmpty)
        {
            throw new ParchmentExtractionException(
                $"Model type '{typeof(TModel).Name}' declares no [EditableField] members — there is nothing to extract.");
        }

        var effectiveCulture = culture ?? CultureInfo.InvariantCulture;

        // The package API needs a seekable stream; buffer non-seekable input.
        var source = docx;
        if (!docx.CanSeek)
        {
            source = new MemoryStream();
            docx.CopyTo(source);
            source.Position = 0;
        }

        try
        {
            using var doc = OpenReadOnly(source);
            var mainPart = doc.MainDocumentPart;
            var body = mainPart?.Document?.Body;
            if (body == null)
            {
                throw new ParchmentExtractionException(
                    "The document has no main body — not a WordprocessingML document.");
            }

            var fields = new List<ExtractedField>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sdt in body.Descendants<SdtElement>())
            {
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (tag == null ||
                    !map.TryGet(tag, out var entry))
                {
                    continue;
                }

                if (!seen.Add(entry.DottedPath))
                {
                    fields.Add(new(entry.DottedPath, FieldState.Duplicate, null, EditableFieldReader.RawText(sdt)));
                    continue;
                }

                fields.Add(EditableFieldReader.Read(sdt, entry, effectiveCulture, mainPart!));
            }

            foreach (var entry in map.Entries)
            {
                if (!seen.Contains(entry.DottedPath))
                {
                    fields.Add(new(entry.DottedPath, FieldState.Missing, null, null));
                }
            }

            return new(fields, map);
        }
        finally
        {
            if (!ReferenceEquals(source, docx))
            {
                source.Dispose();
            }
        }
    }

    static WordprocessingDocument OpenReadOnly(Stream stream)
    {
        try
        {
            return WordprocessingDocument.Open(stream, false);
        }
        catch (Exception exception) when (exception is FileFormatException or InvalidDataException or OpenXmlPackageException)
        {
            throw new ParchmentExtractionException("The stream is not a valid docx package.", exception);
        }
    }
}
