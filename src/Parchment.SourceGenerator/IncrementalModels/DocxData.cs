sealed record DocxData(
    string Path,
    EquatableArray<string> Paragraphs,
    bool HasRemovePersonalInformation,
    string? ReadError);
