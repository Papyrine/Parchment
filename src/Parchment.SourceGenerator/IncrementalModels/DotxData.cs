/// <summary>
/// A style document (<c>.dotx</c>) declared in AdditionalFiles. Consumed by the markdown flow
/// only: a docx template carries its own styles. Bytes travel as base64 — see
/// <see cref="DocxData.ContentBase64"/>.
/// </summary>
sealed record DotxData(
    string Path,
    string ContentBase64,
    string? ReadError);
