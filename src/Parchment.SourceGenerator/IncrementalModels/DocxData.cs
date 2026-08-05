sealed record DocxData(
    string Path,
    EquatableArray<string> Paragraphs,
    EquatableArray<string> BodyParagraphs,
    bool HasRemovePersonalInformation,
    string? ReadError,
    // See MarkdownData.ResourceName.
    string? ResourceName = null);
