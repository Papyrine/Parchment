/// <summary>
/// Rewrites legacy Word FORMTEXT form fields into ordinary <c>{{ Name }}</c> token runs at
/// registration time, so a template authored as a Word form binds like any other docx template.
/// </summary>
/// <remarks>
/// A form field is a run sequence — <c>fldChar begin</c> (carrying the field name in its
/// <c>ffData</c>), the <c>FORMTEXT</c> instruction, <c>fldChar separate</c>, the result runs, then
/// <c>fldChar end</c>. The whole sequence collapses to one run holding the token, keeping the run
/// properties of the field result so the substituted text is formatted as the field was.
/// <para>
/// The runs are taken in document order at any inline depth of the paragraph, so a field inside a
/// hyperlink, a tracked insertion or a content control converts like one directly in the
/// paragraph, and so does one whose runs straddle such a wrapper: the token goes where the
/// <c>begin</c> run was. A field is bounded by its own paragraph — one whose <c>end</c> is in a
/// later paragraph is left as it is.
/// </para>
/// </remarks>
static class FormFields
{
    public static void ToTokens(OpenXmlCompositeElement partRoot)
    {
        // Only the paragraphs that hold a form field are walked; most templates have none.
        var paragraphs = new List<Paragraph>();
        foreach (var data in partRoot.Descendants<FormFieldData>())
        {
            var paragraph = data.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph != null &&
                !paragraphs.Contains(paragraph))
            {
                paragraphs.Add(paragraph);
            }
        }

        foreach (var paragraph in paragraphs)
        {
            Convert(paragraph);
        }
    }

    static void Convert(Paragraph paragraph)
    {
        var runs = new List<Run>();
        CollectRuns(paragraph, runs);
        for (var index = 0; index < runs.Count; index++)
        {
            var name = FieldName(runs[index]);
            if (name == null)
            {
                continue;
            }

            var end = IndexOfEnd(runs, index, out var separate);
            if (end == -1)
            {
                continue;
            }

            var token = new Run(new Text($"{{{{ {name} }}}}")
            {
                Space = SpaceProcessingModeValues.Preserve
            });
            var properties = ResultProperties(runs, separate, end);
            if (properties != null)
            {
                token.RunProperties = properties;
            }

            runs[index].InsertBeforeSelf(token);

            // Only the runs go. Anything else in the span — most often the bookmark Word wraps the
            // field in, or a wrapper some of the field's runs sat in — is left where it is.
            for (var scan = index; scan <= end; scan++)
            {
                runs[scan].Remove();
            }

            index = end;
        }
    }

    // The paragraph's own runs in document order, through hyperlinks, tracked changes, content
    // controls and the like. A run is not descended into, so the paragraphs of a text box anchored
    // in one stay separate; those are converted on their own.
    static void CollectRuns(OpenXmlElement parent, List<Run> runs)
    {
        foreach (var child in parent.ChildElements)
        {
            if (child is Run run)
            {
                runs.Add(run);
                continue;
            }

            if (child.HasChildren)
            {
                CollectRuns(child, runs);
            }
        }
    }

    /// <summary>
    /// The field name, when the run opens a FORMTEXT field whose name can be a token.
    /// </summary>
    static string? FieldName(Run run)
    {
        // A field in a tracked deletion is content the author removed, not a placeholder.
        if (run.Parent is DeletedRun or MoveFromRun)
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

    /// <summary>
    /// The index of the run closing the field opened at <paramref name="start"/>, or -1 when the
    /// paragraph does not close it. Fields nest, so the closing <c>end</c> is the one that balances
    /// the opening <c>begin</c>, not the first one met. <paramref name="separate"/> is the run
    /// holding the field's own separator, or -1.
    /// </summary>
    static int IndexOfEnd(List<Run> runs, int start, out int separate)
    {
        separate = -1;
        var depth = 0;
        for (var index = start; index < runs.Count; index++)
        {
            foreach (var fieldChar in runs[index].Elements<FieldChar>())
            {
                var type = fieldChar.FieldCharType?.Value;
                if (type == FieldCharValues.Begin)
                {
                    depth++;
                    continue;
                }

                if (type == FieldCharValues.Separate)
                {
                    if (depth == 1 &&
                        separate == -1)
                    {
                        separate = index;
                    }

                    continue;
                }

                if (type != FieldCharValues.End)
                {
                    continue;
                }

                depth--;
                if (depth == 0)
                {
                    return index;
                }
            }
        }

        return -1;
    }

    /// <summary>
    /// The run properties of the field result, so the token renders as the field displayed.
    /// </summary>
    static RunProperties? ResultProperties(List<Run> runs, int separate, int end)
    {
        if (separate == -1)
        {
            return null;
        }

        for (var index = separate + 1; index < end; index++)
        {
            if (runs[index].RunProperties is { } properties)
            {
                return (RunProperties) properties.CloneNode(true);
            }
        }

        return null;
    }
}
