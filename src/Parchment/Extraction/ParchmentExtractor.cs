using W15 = DocumentFormat.OpenXml.Office2013.Word;

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
    public static ExtractResult<TModel> Extract<TModel>(string path, CultureInfo? culture = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var file = File.OpenRead(path);
        return Extract<TModel>(file, culture);
    }

    public static ExtractResult<TModel> Extract<TModel>(Stream docx, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(docx);

        // A model the generator never saw yields an empty map too — say that, rather than the
        // misleading "declares no [EditableField] members" a genuinely field-less model earns.
        var map = EditableMap.Build(typeof(TModel));
        if (map.IsEmpty)
        {
            if (!SharedFluid.IsModelRegistered(typeof(TModel)))
            {
                throw new ParchmentExtractionException(SharedFluid.NotGeneratedMessage(typeof(TModel)));
            }

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
            var collections = new List<ExtractedCollection>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenCollections = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Repeating-section pass first: reconstruct each editable collection from its items.
            // The scalar pass below skips every control inside one (their tags are item-relative and
            // non-unique — they must be read grouped by section item, not flat).
            foreach (var sdt in body.Descendants<SdtElement>())
            {
                if (sdt.SdtProperties?.GetFirstChild<W15.SdtRepeatedSection>() == null)
                {
                    continue;
                }

                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (tag == null ||
                    !map.TryGetCollection(tag, out var collectionEntry) ||
                    !seenCollections.Add(collectionEntry.DottedPath))
                {
                    continue;
                }

                collections.Add(new(collectionEntry, FieldState.Extracted, ReconstructCollection(sdt, collectionEntry, effectiveCulture, mainPart!)));
            }

            foreach (var sdt in body.Descendants<SdtElement>())
            {
                // A control inside a repeated item belongs to that row, whichever section holds it:
                // a section this pass did not read (a second one for the same collection, or one
                // for no member of the model) still tags its controls with bare member names, and
                // a root member can share one.
                var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (tag == null ||
                    !map.TryGet(tag, out var entry) ||
                    RepeatedItemOf(sdt) != null)
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

            foreach (var collectionEntry in map.Collections)
            {
                if (!seenCollections.Contains(collectionEntry.DottedPath))
                {
                    collections.Add(new(collectionEntry, FieldState.Missing, null));
                }
            }

            return new(fields, collections, map);
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

    static object ReconstructCollection(SdtElement container, CollectionEntry entry, CultureInfo culture, MainDocumentPart mainPart)
    {
        var listType = typeof(List<>).MakeGenericType(entry.ElementType);
        var list = (IList)Activator.CreateInstance(listType)!;

        foreach (var itemSdt in RepeatedItems(container))
        {
            var item = entry.ElementFactory();
            var isBlank = EditableRepeatingBuilder.IsBlankItem(itemSdt);
            var anyValue = false;
            foreach (var elementEntry in entry.ElementMap.Entries)
            {
                var control = FindControl(itemSdt, elementEntry.DottedPath);
                if (control == null)
                {
                    continue;
                }

                var field = EditableFieldReader.Read(control, elementEntry, culture, mainPart);
                if (field.State == FieldState.Extracted &&
                    elementEntry.CanReach(item))
                {
                    elementEntry.Setter(item, field.Value);

                    // A checkbox cannot be rendered empty, so in the blank item of an empty
                    // collection an unchecked box is still the untouched state, not an entry.
                    if (isBlank &&
                        field.Value is false)
                    {
                        continue;
                    }

                    anyValue = true;
                }
            }

            // Drop all-empty items: the blank clone template Word needs, or a row the user cleared.
            if (anyValue)
            {
                list.Add(item);
            }
        }

        if (!entry.IsArray)
        {
            return list;
        }

        var array = Array.CreateInstance(entry.ElementType, list.Count);
        list.CopyTo(array, 0);
        return array;
    }

    static IEnumerable<SdtElement> RepeatedItems(SdtElement container)
    {
        var content = container.ChildElements.FirstOrDefault(_ => _.LocalName == "sdtContent");
        if (content == null)
        {
            yield break;
        }

        foreach (var child in content.ChildElements)
        {
            if (child is SdtElement sdt &&
                sdt.SdtProperties?.GetFirstChild<W15.SdtRepeatedSectionItem>() != null)
            {
                yield return sdt;
            }
        }
    }

    // The row's own control: one inside a section nested in the row belongs to that section's row.
    static SdtElement? FindControl(SdtElement itemSdt, string tag) =>
        itemSdt.Descendants<SdtElement>()
            .FirstOrDefault(_ => _.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag &&
                                 RepeatedItemOf(_) == itemSdt);

    /// <summary>
    /// The nearest repeated-section item holding <paramref name="sdt"/>, or null for a control
    /// that is in no row.
    /// </summary>
    static SdtElement? RepeatedItemOf(SdtElement sdt)
    {
        foreach (var ancestor in sdt.Ancestors<SdtElement>())
        {
            if (ancestor.SdtProperties?.GetFirstChild<W15.SdtRepeatedSectionItem>() != null)
            {
                return ancestor;
            }
        }

        return null;
    }
}
