/// <summary>
/// Rewrites legacy Word FORMTEXT form fields into ordinary <c>{{ Name }}</c> token runs at
/// registration time, so a template authored as a Word form binds like any other docx template.
/// </summary>
/// <remarks>
/// A form field is a run sequence — <c>fldChar begin</c> (carrying the field name in its
/// <c>ffData</c>), the <c>FORMTEXT</c> instruction, <c>fldChar separate</c>, the result runs, then
/// <c>fldChar end</c>. The whole sequence collapses to one run holding the token, keeping the run
/// properties of the field result so the substituted text is formatted as the field was.
/// </remarks>
static class FormFields
{
    public static void ToTokens(OpenXmlCompositeElement partRoot)
    {
        foreach (var paragraph in partRoot.Descendants<Paragraph>().ToList())
        {
            Convert(paragraph);
        }
    }

    static void Convert(Paragraph paragraph)
    {
        var children = paragraph.ChildElements.ToList();
        for (var index = 0; index < children.Count; index++)
        {
            var name = FieldName(children[index]);
            if (name == null)
            {
                continue;
            }

            var end = IndexOfEnd(children, index);
            if (end == -1)
            {
                continue;
            }

            var token = new Run(new Text($"{{{{ {name} }}}}")
            {
                Space = SpaceProcessingModeValues.Preserve
            });
            var properties = ResultProperties(children, index, end);
            if (properties != null)
            {
                token.RunProperties = properties;
            }

            children[index].InsertBeforeSelf(token);

            // Only the runs go. Anything else in the span — most often the bookmark Word wraps the
            // field in — is left where it is.
            for (var scan = index; scan <= end; scan++)
            {
                if (children[scan] is Run)
                {
                    children[scan].Remove();
                }
            }

            index = end;
        }
    }

    /// <summary>
    /// The field name, when the element opens a FORMTEXT field whose name can be a token.
    /// </summary>
    static string? FieldName(OpenXmlElement element)
    {
        if (element is not Run run)
        {
            return null;
        }

        var fieldChar = run.GetFirstChild<FieldChar>();
        if (fieldChar?.FieldCharType?.Value != FieldCharValues.Begin)
        {
            return null;
        }

        var data = fieldChar.GetFirstChild<FormFieldData>();
        // Checkbox and dropdown fields hold values Word manages, not text to substitute.
        if (data?.GetFirstChild<TextInput>() == null)
        {
            return null;
        }

        var name = data.GetFirstChild<FormFieldName>()?.Val?.Value;
        if (name == null ||
            !IsTokenName(name))
        {
            return null;
        }

        return name;
    }

    // An unnamed field, or one named "Due date", cannot become a member reference — leave it alone
    // rather than emitting a token that could never bind.
    static bool IsTokenName(string name)
    {
        if (name.Length == 0 ||
            !char.IsLetter(name[0]) && name[0] != '_')
        {
            return false;
        }

        foreach (var character in name)
        {
            if (!char.IsLetterOrDigit(character) &&
                character != '_' &&
                character != '.')
            {
                return false;
            }
        }

        return true;
    }

    static int IndexOfEnd(List<OpenXmlElement> children, int start)
    {
        for (var index = start + 1; index < children.Count; index++)
        {
            if (children[index] is Run run &&
                run.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End)
            {
                return index;
            }
        }

        return -1;
    }

    /// <summary>
    /// The run properties of the field result, so the token renders as the field displayed.
    /// </summary>
    static RunProperties? ResultProperties(List<OpenXmlElement> children, int start, int end)
    {
        var separate = -1;
        for (var index = start + 1; index < end; index++)
        {
            if (children[index] is Run run &&
                run.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate)
            {
                separate = index;
                break;
            }
        }

        if (separate == -1)
        {
            return null;
        }

        for (var index = separate + 1; index < end; index++)
        {
            if (children[index] is Run { RunProperties: { } properties })
            {
                return (RunProperties) properties.CloneNode(true);
            }
        }

        return null;
    }
}
