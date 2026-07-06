/// <summary>
/// <paramref name="Paragraphs"/> spans every scanned part (body, headers, footers, notes);
/// <paramref name="BodyParagraphs"/> is the <c>word/document.xml</c> subset. Editable-field
/// token rules are body-scoped (duplicate tags, filter rejection, loop warnings) because the
/// runtime only dispatches editable fields in the body — the same token in a header renders as
/// plain read-only text by design.
/// </summary>
readonly record struct DocxReadResult(
    List<string> Paragraphs,
    List<string> BodyParagraphs,
    bool HasRemovePersonalInformation);
