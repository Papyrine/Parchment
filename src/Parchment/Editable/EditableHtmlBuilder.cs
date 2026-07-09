using SdtLock = DocumentFormat.OpenXml.Wordprocessing.Lock;

/// <summary>
/// Builds the block-level editable-HTML region: <c>[w:permStart, w:sdt(block), w:permEnd]</c>.
/// The block content control (a rich-text <c>w:sdt</c> — no kind element, which is what Word
/// treats as rich text) carries the dotted model path as its <c>w:tag</c> (the round-trip key)
/// and <c>w:lock="sdtLocked"</c> so users can edit the formatted content but not delete the
/// control. The perm range (edGrp="everyone") punches an editable exception through the
/// document's read-only protection around the block — the block-level analogue of
/// <see cref="EditableFieldBuilder"/>. Rendered HTML blocks (from OpenXmlHtml's
/// <c>WordHtmlConverter</c>) are moved into the control's <c>sdtContent</c>; an empty value
/// renders the grey placeholder paragraph.
/// </summary>
static class EditableHtmlBuilder
{
    const string placeholderText = "Click or tap here to enter text.";

    public static IReadOnlyList<OpenXmlElement> Build(
        EditableEntry entry,
        IReadOnlyList<OpenXmlElement> content,
        EditableState state)
    {
        var id = state.NextId();

        var sdtPr = new SdtProperties();
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

        var isPlaceholder = content.Count == 0;
        if (isPlaceholder)
        {
            sdtPr.AppendChild(new ShowingPlaceholder());
        }

        var contentBlock = new SdtContentBlock();
        if (isPlaceholder)
        {
            contentBlock.AppendChild(PlaceholderParagraph());
        }
        else
        {
            foreach (var element in content)
            {
                contentBlock.AppendChild(element);
            }
        }

        return
        [
            new PermStart
            {
                Id = id,
                EditorGroup = RangePermissionEditingGroupValues.Everyone
            },
            new SdtBlock(sdtPr, contentBlock),
            new PermEnd
            {
                Id = id
            }
        ];
    }

    static Paragraph PlaceholderParagraph()
    {
        // Word's built-in grey placeholder look. The style may not exist in the template's styles
        // part — Word tolerates the dangling reference and falls back to plain text.
        var runProperties = new RunProperties(
            new RunStyle
            {
                Val = "PlaceholderText"
            });
        var run = new Run(
            runProperties,
            new Text(placeholderText)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        return new(run);
    }
}
