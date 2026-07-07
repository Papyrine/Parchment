using W14 = DocumentFormat.OpenXml.Office2010.Word;
using SdtLock = DocumentFormat.OpenXml.Wordprocessing.Lock;

/// <summary>
/// Builds the inline element triple for one editable field:
/// <c>[w:permStart, w:sdt, w:permEnd]</c>. The perm range (edGrp="everyone") punches an editable
/// exception through the document's read-only protection; the content control carries the dotted
/// model path as its <c>w:tag</c> (the round-trip key), <c>w:lock="sdtLocked"</c> so users can
/// edit the content but not delete the control, and a kind element per
/// <see cref="EditableFieldKind"/>. Checkbox / date / dropdown controls store canonical values
/// (<c>w14:checked</c>, <c>w:fullDate</c>, <c>w:listItem/@w:value</c>) so extraction does not
/// depend on parsing display text.
/// </summary>
static class EditableFieldBuilder
{
    const string placeholderText = "Click or tap here to enter text.";

    public static IReadOnlyList<OpenXmlElement> Build(
        EditableEntry entry,
        object? value,
        RunProperties? sitePr,
        EditableState state,
        CultureInfo culture)
    {
        var id = state.NextId();

        var sdtPr = new SdtProperties();
        if (sitePr != null)
        {
            sdtPr.AppendChild((RunProperties)sitePr.CloneNode(true));
        }

        sdtPr.AppendChild(new SdtAlias
        {
            Val = entry.Alias
        });
        sdtPr.AppendChild(new Tag
        {
            Val = entry.DottedPath
        });
        sdtPr.AppendChild(new SdtId
        {
            Val = id
        });
        sdtPr.AppendChild(new SdtLock
        {
            Val = LockingValues.SdtLocked
        });

        var (kindElement, content, isPlaceholder) = BuildKind(entry, value, sitePr, culture);
        if (isPlaceholder)
        {
            sdtPr.AppendChild(new ShowingPlaceholder());
        }

        sdtPr.AppendChild(kindElement);

        return
        [
            new PermStart
            {
                Id = id,
                EditorGroup = RangePermissionEditingGroupValues.Everyone
            },
            new SdtRun(sdtPr, new SdtContentRun(content)),
            new PermEnd
            {
                Id = id
            }
        ];
    }

    static (OpenXmlElement KindElement, Run Content, bool IsPlaceholder) BuildKind(
        EditableEntry entry,
        object? value,
        RunProperties? sitePr,
        CultureInfo culture) =>
        entry.Kind switch
        {
            EditableFieldKind.Text => BuildText(entry, value as string, sitePr),
            EditableFieldKind.Number => BuildNumber(value, sitePr, culture),
            EditableFieldKind.Checkbox => BuildCheckbox(value, sitePr),
            EditableFieldKind.Date => BuildDate(entry, value, sitePr, culture),
            EditableFieldKind.DateTimeOffset => BuildTextValue(FormatDateTimeOffset(entry, value, culture), sitePr),
            EditableFieldKind.Time => BuildTextValue(FormatTime(entry, value, culture), sitePr),
            EditableFieldKind.DropDown => BuildDropDown(entry, value, sitePr),
            _ => throw new InvalidOperationException($"Unknown editable kind {entry.Kind}")
        };

    static (OpenXmlElement, Run, bool) BuildText(EditableEntry entry, string? value, RunProperties? sitePr)
    {
        var kind = new SdtContentText();
        if (entry.MultiLine)
        {
            kind.MultiLine = true;
        }

        if (string.IsNullOrEmpty(value))
        {
            return (kind, PlaceholderRun(sitePr), true);
        }

        return (kind, TextRun(value, sitePr, entry.MultiLine), false);
    }

    static (OpenXmlElement, Run, bool) BuildNumber(object? value, RunProperties? sitePr, CultureInfo culture)
    {
        var kind = new SdtContentText();
        if (value == null)
        {
            return (kind, PlaceholderRun(sitePr), true);
        }

        var text = ((IFormattable)value).ToString(null, culture);
        return (kind, TextRun(text, sitePr, multiLine: false), false);
    }

    static (OpenXmlElement, Run, bool) BuildCheckbox(object? value, RunProperties? sitePr)
    {
        // bool? is rejected at map build, so value is always a non-null bool here.
        var isChecked = (bool)value!;
        var kind = new W14.SdtContentCheckBox(
            new W14.Checked
            {
                Val = isChecked ? W14.OnOffValues.One : W14.OnOffValues.Zero
            },
            new W14.CheckedState
            {
                Font = "MS Gothic",
                Val = "2612"
            },
            new W14.UncheckedState
            {
                Font = "MS Gothic",
                Val = "2610"
            });

        // The glyph must render from a font that has it; force MS Gothic on the content run,
        // matching what Word itself emits when a checkbox is toggled.
        var pr = ClonePr(sitePr);
        pr.RemoveAllChildren<RunFonts>();
        InsertAfterRunStyle(
            pr,
            new RunFonts
            {
                Ascii = "MS Gothic",
                HighAnsi = "MS Gothic",
                EastAsia = "MS Gothic"
            });

        var run = new Run(pr);
        run.AppendChild(new Text(isChecked ? "☒" : "☐"));
        return (kind, run, false);
    }

    static (OpenXmlElement, Run, bool) BuildDate(EditableEntry entry, object? value, RunProperties? sitePr, CultureInfo culture)
    {
        var format = entry.DateFormat ?? "yyyy-MM-dd";
        var kind = new SdtContentDate(
            new DateFormat
            {
                Val = format
            });

        if (value == null)
        {
            return (kind, PlaceholderRun(sitePr), true);
        }

        // Only DateOnly and DateTime reach here (see BuildKind). DateTime's time-of-day is
        // preserved in w:fullDate even though the default format shows date only.
        var dateTime = value switch
        {
            Date date => date.ToDateTime(Time.MinValue),
            _ => (DateTime)value
        };
        kind.FullDate = dateTime;
        return (kind, TextRun(dateTime.ToString(format, culture), sitePr, multiLine: false), false);
    }

    // Round-trippable ISO 8601 defaults for the text-backed temporal kinds. DateTimeOffset keeps
    // its offset (zzz); both are re-parsed verbatim by EditableFieldReader. Seconds precision by
    // default — a caller needing sub-seconds sets DateFormat (e.g. "o").
    const string dateTimeOffsetFormat = "yyyy-MM-ddTHH:mm:sszzz";
    const string timeFormat = "HH:mm:ss";

    static string? FormatDateTimeOffset(EditableEntry entry, object? value, CultureInfo culture) =>
        value is DateTimeOffset offset
            ? offset.ToString(entry.DateFormat ?? dateTimeOffsetFormat, culture)
            : null;

    static string? FormatTime(EditableEntry entry, object? value, CultureInfo culture) =>
        value is Time time
            ? time.ToString(entry.DateFormat ?? timeFormat, culture)
            : null;

    static (OpenXmlElement, Run, bool) BuildTextValue(string? text, RunProperties? sitePr)
    {
        var kind = new SdtContentText();
        if (string.IsNullOrEmpty(text))
        {
            return (kind, PlaceholderRun(sitePr), true);
        }

        return (kind, TextRun(text, sitePr, multiLine: false), false);
    }

    static (OpenXmlElement, Run, bool) BuildDropDown(EditableEntry entry, object? value, RunProperties? sitePr)
    {
        var kind = new SdtContentDropDownList();
        foreach (var name in Enum.GetNames(entry.ClrType))
        {
            kind.AppendChild(
                new ListItem
                {
                    DisplayText = name,
                    Value = name
                });
        }

        if (value == null)
        {
            return (kind, PlaceholderRun(sitePr), true);
        }

        return (kind, TextRun(value.ToString()!, sitePr, multiLine: false), false);
    }

    static Run TextRun(string value, RunProperties? sitePr, bool multiLine)
    {
        var run = new Run();
        if (sitePr != null)
        {
            run.AppendChild((RunProperties)sitePr.CloneNode(true));
        }

        var cleaned = XmlCharSanitizer.Strip(value).ToString();
        if (multiLine)
        {
            var first = true;
            foreach (var line in cleaned.Split('\n'))
            {
                if (!first)
                {
                    run.AppendChild(new Break());
                }

                first = false;
                if (line.Length > 0)
                {
                    run.AppendChild(
                        new Text(line.TrimEnd('\r'))
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        });
                }
            }
        }
        else
        {
            run.AppendChild(
                new Text(cleaned.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' '))
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
        }

        return run;
    }

    static Run PlaceholderRun(RunProperties? sitePr)
    {
        var pr = ClonePr(sitePr);
        // Word's built-in grey placeholder look. The style may not exist in the template's
        // styles part — Word tolerates the dangling reference and falls back to plain text.
        pr.InsertAt(
            new RunStyle
            {
                Val = "PlaceholderText"
            },
            0);
        var run = new Run(pr);
        run.AppendChild(
            new Text(placeholderText)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        return run;
    }

    static RunProperties ClonePr(RunProperties? sitePr) =>
        sitePr == null ? new() : (RunProperties)sitePr.CloneNode(true);

    static void InsertAfterRunStyle(RunProperties pr, OpenXmlElement element)
    {
        // rPr schema order: rStyle first, rFonts immediately after.
        if (pr.GetFirstChild<RunStyle>() is { } style)
        {
            pr.InsertAfter(element, style);
        }
        else
        {
            pr.InsertAt(element, 0);
        }
    }
}
