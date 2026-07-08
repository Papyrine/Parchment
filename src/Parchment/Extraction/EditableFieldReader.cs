using W14 = DocumentFormat.OpenXml.Office2010.Word;

/// <summary>
/// Reads one editable content control back into a typed value. Checkbox / date / dropdown
/// controls carry canonical values in their <c>sdtPr</c> (<c>w14:checked</c>, <c>w:fullDate</c>,
/// <c>w:listItem/@w:value</c>), so those never depend on parsing display text. Numerics, and the
/// text-backed temporal kinds (<c>DateTimeOffset</c>, <c>TimeOnly</c>, which have no canonical
/// control slot), parse the run text with the caller's culture — which must match the render
/// culture.
/// </summary>
static class EditableFieldReader
{
    public static ExtractedField Read(SdtElement sdt, EditableEntry entry, CultureInfo culture)
    {
        var raw = RawText(sdt);
        if (entry.Kind == EditableFieldKind.Checkbox)
        {
            return ReadCheckbox(sdt, entry, raw);
        }

        var isPlaceholder = sdt.SdtProperties?.GetFirstChild<ShowingPlaceholder>() != null;
        return entry.Kind switch
        {
            EditableFieldKind.Text when isPlaceholder || raw.Length == 0 =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.Text =>
                new(entry.DottedPath, FieldState.Extracted, raw, raw),

            EditableFieldKind.Number when isPlaceholder || raw.Length == 0 =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.Number =>
                ReadNumber(entry, raw, culture),

            EditableFieldKind.Date when isPlaceholder =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.Date =>
                ReadDate(sdt, entry, raw, culture),

            EditableFieldKind.DateTimeOffset when isPlaceholder || raw.Length == 0 =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.DateTimeOffset =>
                ReadDateTimeOffset(entry, raw, culture),

            EditableFieldKind.Time when isPlaceholder || raw.Length == 0 =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.Time =>
                ReadTime(entry, raw, culture),

            EditableFieldKind.DropDown when isPlaceholder || raw.Length == 0 =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.DropDown =>
                ReadDropDown(sdt, entry, raw),

            _ => throw new InvalidOperationException($"Unknown editable kind {entry.Kind}")
        };
    }

    public static string RawText(SdtElement sdt)
    {
        var content = sdt.ChildElements.FirstOrDefault(_ => _.LocalName == "sdtContent");
        if (content == null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        foreach (var element in content.Descendants())
        {
            if (element is Text text)
            {
                builder.Append(text.Text);
            }
            else if (element is Break)
            {
                builder.Append('\n');
            }
        }

        return builder.ToString();
    }

    static ExtractedField ReadCheckbox(SdtElement sdt, EditableEntry entry, string raw)
    {
        var isChecked = sdt.SdtProperties
            ?.GetFirstChild<W14.SdtContentCheckBox>()
            ?.GetFirstChild<W14.Checked>()
            ?.Val;
        if (isChecked is { HasValue: true })
        {
            return new(entry.DottedPath, FieldState.Extracted, isChecked.Value == W14.OnOffValues.One, raw);
        }

        // Fall back to the rendered glyph when w14:checked is absent.
        return raw switch
        {
            "☒" => new(entry.DottedPath, FieldState.Extracted, true, raw),
            "☐" => new(entry.DottedPath, FieldState.Extracted, false, raw),
            _ => new(entry.DottedPath, FieldState.ParseFailed, null, raw)
        };
    }

    static ExtractedField ReadNumber(EditableEntry entry, string raw, CultureInfo culture)
    {
        if (TryParseNumber(raw, entry.ClrType, culture, out var value))
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
    }

    static bool TryParseNumber(string raw, Type type, CultureInfo culture, out object? value)
    {
        value = Type.GetTypeCode(type) switch
        {
            TypeCode.Byte when byte.TryParse(raw, NumberStyles.Integer, culture, out var b) => b,
            TypeCode.SByte when sbyte.TryParse(raw, NumberStyles.Integer, culture, out var sb) => sb,
            TypeCode.Int16 when short.TryParse(raw, NumberStyles.Integer, culture, out var s) => s,
            TypeCode.UInt16 when ushort.TryParse(raw, NumberStyles.Integer, culture, out var us) => us,
            TypeCode.Int32 when int.TryParse(raw, NumberStyles.Integer, culture, out var i) => i,
            TypeCode.UInt32 when uint.TryParse(raw, NumberStyles.Integer, culture, out var ui) => ui,
            TypeCode.Int64 when long.TryParse(raw, NumberStyles.Integer, culture, out var l) => l,
            TypeCode.UInt64 when ulong.TryParse(raw, NumberStyles.Integer, culture, out var ul) => ul,
            TypeCode.Single when float.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, culture, out var f) => f,
            TypeCode.Double when double.TryParse(raw, NumberStyles.Float | NumberStyles.AllowThousands, culture, out var d) => d,
            TypeCode.Decimal when decimal.TryParse(raw, NumberStyles.Number, culture, out var m) => m,
            _ => null
        };
        return value != null;
    }

    static ExtractedField ReadDate(SdtElement sdt, EditableEntry entry, string raw, CultureInfo culture)
    {
        var fullDate = sdt.SdtProperties
            ?.GetFirstChild<SdtContentDate>()
            ?.FullDate;
        DateTime dateTime;
        if (fullDate is { HasValue: true })
        {
            dateTime = fullDate.Value;
        }
        else if (raw.Length == 0)
        {
            return new(entry.DottedPath, FieldState.Empty, null, raw);
        }
        else
        {
            // Word maintains w:fullDate for picker selections; typed-in text may leave it stale
            // or absent, so parse the display text with the declared format.
            var format = entry.DateFormat ?? "yyyy-MM-dd";
            if (!DateTime.TryParseExact(raw, format, culture, DateTimeStyles.None, out dateTime) &&
                !DateTime.TryParse(raw, culture, DateTimeStyles.None, out dateTime))
            {
                return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
            }
        }

        // Only DateOnly and DateTime map to the Date kind. DateTime comes back with Kind
        // Unspecified (w:fullDate carries no zone) — documented; callers needing a specific Kind
        // re-stamp it.
        object value = entry.ClrType == typeof(Date)
            ? Date.FromDateTime(dateTime)
            : dateTime;

        return new(entry.DottedPath, FieldState.Extracted, value, raw);
    }

    static ExtractedField ReadDateTimeOffset(EditableEntry entry, string raw, CultureInfo culture)
    {
        // The run text is the source of truth (w:fullDate cannot hold an offset). Parsing the
        // offset-bearing ISO text preserves the original offset — DateTimeOffset.TryParse keeps
        // whatever offset the string carries.
        var parsed = entry.DateFormat is { } format
            ? DateTimeOffset.TryParseExact(raw, format, culture, DateTimeStyles.None, out var value)
            : DateTimeOffset.TryParse(raw, culture, DateTimeStyles.None, out value);

        if (parsed)
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
    }

    static ExtractedField ReadTime(EditableEntry entry, string raw, CultureInfo culture)
    {
        var parsed = entry.DateFormat is { } format
            ? Time.TryParseExact(raw, format, culture, DateTimeStyles.None, out var value)
            : Time.TryParse(raw, culture, DateTimeStyles.None, out value);

        if (parsed)
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
    }

    static ExtractedField ReadDropDown(SdtElement sdt, EditableEntry entry, string raw)
    {
        // Display text → w:listItem/@w:val, falling back to the text itself (our controls emit
        // displayText == value == enum member name, but a hand-edited doc may differ).
        var candidate = sdt.SdtProperties
                            ?.GetFirstChild<SdtContentDropDownList>()
                            ?.Elements<ListItem>()
                            .FirstOrDefault(_ => _.DisplayText?.Value == raw)
                            ?.Value?.Value ??
                        raw;

        if (Enum.TryParse(entry.ClrType, candidate, ignoreCase: true, out var parsed) &&
            Enum.IsDefined(entry.ClrType, parsed))
        {
            return new(entry.DottedPath, FieldState.Extracted, parsed, raw);
        }

        return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
    }
}
