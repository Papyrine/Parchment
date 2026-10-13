using W14 = DocumentFormat.OpenXml.Office2010.Word;

/// <summary>
/// Reads one editable content control back into a typed value. Checkbox / date / dropdown
/// controls carry canonical values in their <c>sdtPr</c> (<c>w14:checked</c>, <c>w:fullDate</c>,
/// <c>w:listItem/@w:value</c>), so those never depend on parsing display text. Numerics, and the
/// text-backed temporal kinds (<c>DateTimeOffset</c>, <c>TimeOnly</c>, which have no canonical
/// control slot), parse the run text. The render always writes that text with the invariant
/// culture, so every parse tries the invariant form first — an unedited control reads back as
/// rendered whatever culture the caller passes — and the caller's culture only widens what a user
/// may type instead.
/// </summary>
static class EditableFieldReader
{
    public static ExtractedField Read(SdtElement sdt, EditableEntry entry, CultureInfo culture, MainDocumentPart mainPart)
    {
        var raw = RawText(sdt);
        if (entry.Kind == EditableFieldKind.Checkbox)
        {
            return ReadCheckbox(sdt, entry, raw);
        }

        // w:showingPlcHdr is an on/off element: w:val="0" means the placeholder is not showing.
        var isPlaceholder = OpenXmlHtmlSerializer.IsOn(sdt.SdtProperties?.GetFirstChild<ShowingPlaceholder>());
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

            EditableFieldKind.Html when isPlaceholder =>
                new(entry.DottedPath, FieldState.Empty, null, raw),
            EditableFieldKind.Html =>
                ReadHtml(sdt, entry, mainPart),

            _ => throw new InvalidOperationException($"Unknown editable kind {entry.Kind}")
        };
    }

    static ExtractedField ReadHtml(SdtElement sdt, EditableEntry entry, MainDocumentPart mainPart)
    {
        var content = sdt.ChildElements.FirstOrDefault(_ => _.LocalName == "sdtContent");
        if (content == null)
        {
            return new(entry.DottedPath, FieldState.Empty, null, string.Empty);
        }

        // The (possibly Word-edited) block content is the source of truth — serialize it back to
        // an HTML string. Empty content extracts as Empty so ApplyTo can null a nullable member.
        var html = OpenXmlHtmlSerializer.ToHtml(content, mainPart);
        if (html.Length == 0)
        {
            return new(entry.DottedPath, FieldState.Empty, null, string.Empty);
        }

        return new(entry.DottedPath, FieldState.Extracted, html, html);
    }

    public static string RawText(SdtElement sdt)
    {
        var content = sdt.ChildElements.FirstOrDefault(_ => _.LocalName == "sdtContent");
        if (content == null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        foreach (var piece in EditableContent.Runs(content))
        {
            if (piece is not Run)
            {
                builder.Append(piece.InnerText);
                continue;
            }

            foreach (var element in piece.Elements())
            {
                builder.Append(RunCharacters.For(element));
            }
        }

        return builder.ToString();
    }

    static ExtractedField ReadCheckbox(SdtElement sdt, EditableEntry entry, string raw)
    {
        // Without w14:checkbox the control is no longer a checkbox. Its glyph is configurable
        // display text, so the state is reported as unreadable rather than guessed from it.
        var checkBox = sdt.SdtProperties?.GetFirstChild<W14.SdtContentCheckBox>();
        if (checkBox == null)
        {
            return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
        }

        // w14:checked is optional; a checkbox without it is unchecked.
        var state = checkBox.GetFirstChild<W14.Checked>();
        if (state == null)
        {
            return new(entry.DottedPath, FieldState.Extracted, false, raw);
        }

        // An on/off element with no w14:val is on.
        var val = state.Val;
        if (val == null)
        {
            return new(entry.DottedPath, FieldState.Extracted, true, raw);
        }

        // w14:ST_OnOff has two spellings of each state: "1" / "true" and "0" / "false".
        if (!val.HasValue)
        {
            return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
        }

        var value = val.Value;
        var isChecked = value == W14.OnOffValues.One || value == W14.OnOffValues.True;
        return new(entry.DottedPath, FieldState.Extracted, isChecked, raw);
    }

    static ExtractedField ReadNumber(EditableEntry entry, string raw, CultureInfo culture)
    {
        if (TryParseNumber(raw, entry.ClrType, culture, out var value))
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
    }

    // Group separators are never accepted: one culture's group separator is another's decimal
    // separator, so "12,5" read invariantly would silently become 125, and so would the rendered
    // "12.5" read as de-DE. Without them a text has at most one reading — the invariant one the
    // render wrote ('.' is always the decimal point), else the caller's culture, which adds its
    // own decimal separator. Anything grouped ("1,000", "1.234,56") is a parse failure, as are
    // currency and percent symbols.
    const NumberStyles decimalStyle = NumberStyles.Number & ~NumberStyles.AllowThousands;

    static bool TryParseNumber(string raw, Type type, CultureInfo culture, out object? value)
    {
        if (TryParseNumberAs(raw, type, CultureInfo.InvariantCulture, out value))
        {
            return true;
        }

        return TryParseNumberAs(raw, type, culture, out value);
    }

    static bool TryParseNumberAs(string raw, Type type, CultureInfo culture, out object? value)
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
            TypeCode.Single when float.TryParse(raw, NumberStyles.Float, culture, out var f) => f,
            TypeCode.Double when double.TryParse(raw, NumberStyles.Float, culture, out var d) => d,
            TypeCode.Decimal when decimal.TryParse(raw, decimalStyle, culture, out var m) => m,
            _ => null
        };
        return value != null;
    }

    static ExtractedField ReadDate(SdtElement sdt, EditableEntry entry, string raw, CultureInfo culture)
    {
        if (!TryReadFullDate(sdt, out var dateTime))
        {
            if (raw.Length == 0)
            {
                return new(entry.DottedPath, FieldState.Empty, null, raw);
            }

            // Word maintains w:fullDate for picker selections; typed-in text may leave it stale
            // or absent, so parse the display text with the declared format, as the render wrote
            // it (invariant) first. The wall clock as typed is kept: an offset in the text is not
            // converted to the machine's zone.
            var format = entry.DateFormat ?? "yyyy-MM-dd";
            if (!TryParseExactDated(raw, format, CultureInfo.InvariantCulture, out var typed) &&
                !TryParseExactDated(raw, format, culture, out typed) &&
                !TryParseDated(raw, culture, out typed))
            {
                return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
            }

            dateTime = typed.DateTime;
        }

        // Only DateOnly and DateTime map to the Date kind. DateTime comes back with Kind
        // Unspecified (the value is a wall-clock reading) — documented; callers needing a
        // specific Kind re-stamp it.
        object value = entry.ClrType == typeof(Date)
            ? Date.FromDateTime(dateTime)
            : DateTime.SpecifyKind(dateTime, DateTimeKind.Unspecified);

        return new(entry.DottedPath, FieldState.Extracted, value, raw);
    }

    // Reads the attribute text rather than the typed value: the typed value converts a "Z" or
    // offset suffix (Word's picker writes "Z") into a Utc DateTime, shifting the clock reading
    // when an offset is present. DateTimeOffset.DateTime is the clock reading as written,
    // whatever suffix follows it; AssumeUniversal keeps a suffix-less value off the machine zone.
    static bool TryReadFullDate(SdtElement sdt, out DateTime dateTime)
    {
        var text = sdt.SdtProperties
            ?.GetFirstChild<SdtContentDate>()
            ?.FullDate
            ?.InnerText;
        if (text != null &&
            DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
        {
            dateTime = parsed.DateTime;
            return true;
        }

        dateTime = default;
        return false;
    }

    static ExtractedField ReadDateTimeOffset(EditableEntry entry, string raw, CultureInfo culture)
    {
        // The run text is the source of truth (w:fullDate cannot hold an offset). Parsing the
        // offset-bearing ISO text preserves the original offset — DateTimeOffset.TryParse keeps
        // whatever offset the string carries. The rendered (invariant) form is tried first: the
        // caller's culture may name months differently or default to another calendar, which
        // would fail on, or shift the year of, text nobody edited. Text without an offset reads
        // as +00:00.
        var renderFormat = entry.DateFormat ?? EditableFieldBuilder.DateTimeOffsetFormat;
        if (TryParseExactDated(raw, renderFormat, CultureInfo.InvariantCulture, out var value))
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        bool parsed;
        if (entry.DateFormat is { } format)
        {
            parsed = TryParseExactDated(raw, format, culture, out value);
        }
        else
        {
            parsed = TryParseDated(raw, culture, out value);
        }

        if (parsed)
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        return new(entry.DottedPath, FieldState.ParseFailed, null, raw);
    }

    // The framework parser fills whatever the text leaves out from the machine: a missing offset
    // becomes the local zone's, a missing date becomes today, and a missing year becomes the
    // current one. Extraction must return the same value on any machine on any day, so a missing
    // offset is fixed at zero (AssumeUniversal) and text that does not state its year is a parse
    // failure. A missing day or month defaults to 1, which is already stable.
    const DateTimeStyles datedStyles = DateTimeStyles.AssumeUniversal;

    static bool TryParseExactDated(string raw, string format, CultureInfo culture, out DateTimeOffset value)
    {
        if (FormatHasYear(format))
        {
            return DateTimeOffset.TryParseExact(raw, format, culture, datedStyles, out value);
        }

        value = default;
        return false;
    }

    // Free-form text has no format to say whether a year was read, so the parsed year has to
    // appear in the text as a four digit number. That also rejects two digit years, whose
    // century pivot is an operating system setting.
    static bool TryParseDated(string raw, CultureInfo culture, out DateTimeOffset value) =>
        DateTimeOffset.TryParse(raw, culture, datedStyles, out value) &&
        HasYear(raw, culture.Calendar.GetYear(value.DateTime));

    // The year-less standard formats are month/day ("M", "m") and the two time patterns.
    static bool FormatHasYear(string format)
    {
        if (format.Length == 1)
        {
            return format is not ("M" or "m" or "t" or "T");
        }

        return format.Contains('y');
    }

    static bool HasYear(string raw, int year)
    {
        var digits = year.ToString("0000", CultureInfo.InvariantCulture);
        var index = raw.IndexOf(digits, StringComparison.Ordinal);
        while (index >= 0)
        {
            var end = index + digits.Length;
            if ((index == 0 || !char.IsAsciiDigit(raw[index - 1])) &&
                (end == raw.Length || !char.IsAsciiDigit(raw[end])))
            {
                return true;
            }

            index = raw.IndexOf(digits, index + 1, StringComparison.Ordinal);
        }

        return false;
    }

    static ExtractedField ReadTime(EditableEntry entry, string raw, CultureInfo culture)
    {
        var renderFormat = entry.DateFormat ?? EditableFieldBuilder.TimeFormat;
        if (Time.TryParseExact(raw, renderFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var value))
        {
            return new(entry.DottedPath, FieldState.Extracted, value, raw);
        }

        var parsed = entry.DateFormat is { } format
            ? Time.TryParseExact(raw, format, culture, DateTimeStyles.None, out value)
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
