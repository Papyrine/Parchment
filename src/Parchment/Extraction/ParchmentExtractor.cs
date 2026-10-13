using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace Parchment;

/// <summary>
/// Reads <c>[EditableField]</c> values back out of a docx produced by a Parchment template —
/// the other half of two-way binding. Content controls are matched by their <c>w:tag</c>
/// (the dotted model path), so no <see cref="TemplateStore"/> or template registration is
/// required. Checkbox / date / dropdown controls are read from canonical control state; string
/// and numeric fields read display text. That text is always rendered with the invariant culture,
/// and is read back that way first, so an unedited document extracts the rendered values whatever
/// <c>culture</c> is passed. The <c>culture</c> argument (default
/// <see cref="CultureInfo.InvariantCulture"/>) describes what a user may type instead: its decimal
/// separator and date forms. Group separators are never accepted in a number.
/// </summary>
public static class ParchmentExtractor
{
    [RequiresDynamicCode(Unsupported.DynamicCode)]
    public static ExtractResult<TModel> Extract<TModel>(string path, CultureInfo? culture = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var file = File.OpenRead(path);
        return Extract<TModel>(file, culture);
    }

    [RequiresDynamicCode(Unsupported.DynamicCode)]
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
            var rowFields = new List<ExtractedField>();
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

                collections.Add(new(collectionEntry, FieldState.Extracted, ReconstructCollection(sdt, collectionEntry, effectiveCulture, mainPart!, rowFields)));
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

            fields.AddRange(rowFields);

            return new(fields, collections, map);
        }
        // The SDK reads each part on first touch, so a package that opened cleanly can still fail
        // here: at the main part, or at a numbering / hyperlink lookup while reading a field.
        catch (XmlException exception)
        {
            throw new ParchmentExtractionException(
                $"The document contains a part that is not well-formed XML: {exception.Message}",
                exception);
        }
        catch (Exception exception) when (exception is FileFormatException or InvalidDataException or OpenXmlPackageException)
        {
            throw new ParchmentExtractionException(
                $"The document could not be read: {exception.GetBaseException().Message}",
                exception);
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

    // Reads the rows as values rather than building the items here: ExtractedCollection.Build
    // constructs the list and its items once per ApplyTo, so no two models share them.
    static List<List<(EditableEntry Entry, object? Value)>> ReconstructCollection(SdtElement container, CollectionEntry entry, CultureInfo culture, MainDocumentPart mainPart, List<ExtractedField> rowFields)
    {
        var rows = new List<List<(EditableEntry Entry, object? Value)>>();

        foreach (var itemSdt in RepeatedItems(container))
        {
            // Never handed out: only answers CanReach for an element as the factory builds it.
            var item = entry.ElementFactory();
            var isBlank = EditableRepeatingBuilder.IsBlankItem(itemSdt);
            var anyValue = false;
            var anyUnchecked = false;
            var resetByWord = false;
            var row = new List<(EditableEntry Entry, object? Value)>();
            var itemFields = new List<ExtractedField>();
            foreach (var elementEntry in entry.ElementMap.Entries)
            {
                var control = FindControl(itemSdt, elementEntry.DottedPath);
                if (control == null)
                {
                    itemFields.Add(new(elementEntry.DottedPath, FieldState.Missing, null, null));
                    continue;
                }

                if (IsResetByWord(control, elementEntry))
                {
                    resetByWord = true;
                }

                var field = EditableFieldReader.Read(control, elementEntry, culture, mainPart);

                // Text that does not parse is still something the user entered: the row is kept
                // and the failure reported, with the member left at the element's default.
                if (field.State == FieldState.ParseFailed)
                {
                    anyValue = true;
                }

                if (field.State != FieldState.Extracted)
                {
                    itemFields.Add(field);
                    continue;
                }

                // The item is new, so an object between it and the member is null unless the
                // element initializes it. The setter creates the ones it can.
                elementEntry.Setter(item, field.Value);

                // Where an object on the way could not be created the value has nowhere to go,
                // and the row's entry says so instead of the value being dropped.
                if (elementEntry.CanReach(item))
                {
                    row.Add((elementEntry, field.Value));
                }
                else
                {
                    field = field with
                    {
                        State = FieldState.Unreachable
                    };
                }

                itemFields.Add(field);

                // A checkbox cannot be rendered empty, so an unchecked box is a value only
                // where the row is known to be one rendered from an item (decided below).
                if (field.Value is false)
                {
                    anyUnchecked = true;
                    continue;
                }

                anyValue = true;
            }

            // Drop all-empty items: the blank clone template Word needs, a row the user cleared, or
            // a row the user added and left alone. With nothing but unchecked boxes to go on, the
            // row is an entry unless it is the blank item of an empty collection (or a copy of it)
            // or holds a control Word reset, which a row rendered from an item never does. A
            // dropped row reports nothing, so a row's index is its position in the rebuilt list.
            var uncheckedEntry = anyUnchecked && !isBlank && !resetByWord;
            if (!anyValue &&
                !uncheckedEntry)
            {
                continue;
            }

            foreach (var field in itemFields)
            {
                rowFields.Add(field with { Path = $"{entry.DottedPath}[{rows.Count}].{field.Path}" });
            }

            rows.Add(row);
        }

        return rows;
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

    // Word's "+" copies a row and resets every control that held a value to a placeholder of
    // spaces, and clearing a control leaves the same. The render never writes that: its own
    // placeholders are words ("Enter text"), which Word keeps when it copies a row showing them.
    static bool IsResetByWord(SdtElement control, EditableEntry entry) =>
        entry.Kind != EditableFieldKind.Checkbox &&
        control.SdtProperties?.GetFirstChild<ShowingPlaceholder>() != null &&
        string.IsNullOrWhiteSpace(EditableFieldReader.RawText(control));

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
