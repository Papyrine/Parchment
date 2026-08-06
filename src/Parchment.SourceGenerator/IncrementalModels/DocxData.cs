sealed record DocxData(
    string Path,
    EquatableArray<string> Paragraphs,
    EquatableArray<string> BodyParagraphs,
    bool HasRemovePersonalInformation,
    // The template's raw bytes, carried as base64 because that is both the literal the emission
    // writes and a cheaply-equatable string for the incremental cache. It has to travel with the
    // record: a change that leaves every paragraph intact (a style tweak) still has to re-emit
    // the embedded copy.
    string ContentBase64,
    string? ReadError);
