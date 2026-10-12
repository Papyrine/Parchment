sealed record DocxData(
    string Path,
    EquatableArray<string> Paragraphs,
    EquatableArray<string> BodyParagraphs,
    // One id per entry of Paragraphs, equal for paragraphs that share a parent element. An int
    // rather than anything that points back into the XDocument, so the record stays equatable.
    EquatableArray<int> ParagraphContainers,
    // One part index per entry of Paragraphs. The runtime builds a scope tree per part.
    EquatableArray<int> ParagraphParts,
    bool HasRemovePersonalInformation,
    // The template's raw bytes, carried as base64 because that is both the literal the emission
    // writes and a cheaply-equatable string for the incremental cache. It has to travel with the
    // record: a change that leaves every paragraph intact (a style tweak) still has to re-emit
    // the embedded copy.
    string ContentBase64,
    string? ReadError);
