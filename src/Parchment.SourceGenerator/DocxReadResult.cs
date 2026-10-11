/// <summary>
/// <paramref name="Paragraphs"/> spans every scanned part (body, headers, footers, notes);
/// <paramref name="BodyParagraphs"/> is the <c>word/document.xml</c> subset. Editable-field
/// token rules are body-scoped (duplicate tags, filter rejection, loop warnings) because the
/// runtime only dispatches editable fields in the body — the same token in a header renders as
/// plain read-only text by design. <paramref name="ParagraphContainers"/> runs parallel to
/// <paramref name="Paragraphs"/>: two paragraphs carry the same id exactly when they share a parent
/// element, which is what the block-tag sibling rule (PARCH028) compares.
/// </summary>
readonly record struct DocxReadResult(
    List<string> Paragraphs,
    List<string> BodyParagraphs,
    List<int> ParagraphContainers,
    bool HasRemovePersonalInformation);
