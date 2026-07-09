using W15 = DocumentFormat.OpenXml.Office2013.Word;

/// <summary>
/// Builds the Word repeating-section structure for an editable collection: a block <c>w:sdt</c>
/// carrying <c>w15:repeatingSection</c> (tagged with the collection's dotted path) wrapping one
/// <c>w15:repeatingSectionItem</c> per model item, all inside a <c>permStart</c>/<c>permEnd</c>
/// editable range so Word lets users add / remove / reorder rows under the document's read-only
/// protection (validated in Word — see the Phase 2 plan). Each item's inner controls are ordinary
/// <see cref="EditableFieldBuilder"/> / <see cref="EditableHtmlBuilder"/> controls tagged with the
/// item-relative path.
/// </summary>
static class EditableRepeatingBuilder
{
    public static SdtBlock BuildItem(IReadOnlyList<OpenXmlElement> body, EditableState state)
    {
        var content = new SdtContentBlock();
        foreach (var element in body)
        {
            // Inner controls arrive wrapped in their own perm range (from EditableFieldBuilder /
            // EditableHtmlBuilder). Inside the repeating section the container's perm range already
            // makes everything editable, and per-item perm markers are duplicated (with colliding
            // ids) when Word clones a row — so strip them, keeping only the controls.
            if (element is PermStart or PermEnd)
            {
                continue;
            }

            foreach (var perm in element.Descendants<PermStart>().ToList())
            {
                perm.Remove();
            }

            foreach (var perm in element.Descendants<PermEnd>().ToList())
            {
                perm.Remove();
            }

            content.AppendChild(element);
        }

        return new(
            new SdtProperties(
                new SdtId
                {
                    Val = state.NextId()
                },
                new W15.SdtRepeatedSectionItem()),
            content);
    }

    public static IReadOnlyList<OpenXmlElement> BuildContainer(CollectionEntry entry, IReadOnlyList<SdtBlock> items, EditableState state)
    {
        var id = state.NextId();

        var content = new SdtContentBlock();
        foreach (var item in items)
        {
            content.AppendChild(item);
        }

        // No sdtLocked on the container or items: Word's add/remove operates on the items, and the
        // spike confirmed the perm range alone gives the desired editable-but-otherwise-locked doc.
        var container = new SdtBlock(
            new SdtProperties(
                new SdtAlias
                {
                    Val = entry.Alias
                },
                new Tag
                {
                    Val = entry.DottedPath
                },
                new SdtId
                {
                    Val = id
                },
                new W15.SdtRepeatedSection()),
            content);

        return
        [
            new PermStart
            {
                Id = id,
                EditorGroup = RangePermissionEditingGroupValues.Everyone
            },
            container,
            new PermEnd
            {
                Id = id
            }
        ];
    }
}
