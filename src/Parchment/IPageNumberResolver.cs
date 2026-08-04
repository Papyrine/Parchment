namespace Parchment;

/// <summary>
/// Supplies the page each bookmark falls on, so a table of contents and page references can be
/// resolved when the document is built rather than when Word next opens it.
/// </summary>
/// <remarks>
/// A page number is a product of layout, which Parchment does not do — it writes OpenXML, it does
/// not paginate it. Ordinarily that leaves the fields for Word to compute, which costs the reader a
/// prompt on open and shows placeholder text until they accept it. Supplying a resolver trades that
/// for a rendering pass: whatever can paginate the document reports where each bookmark landed, the
/// numbers are written into the fields' cached results, and Word opens the document quietly with the
/// values already in place.
/// <para>
/// The fields themselves survive, so a reader can still refresh them. That matters, because a
/// resolver is only as right as its agreement with Word: a renderer that breaks a page one line
/// earlier will report numbers that are wrong and, with nothing marked dirty, look authoritative.
/// Prefer a resolver whose layout is known to track Word's for the documents in question.
/// </para>
/// <para>
/// Called once per render with the fully built document, positioned at its start. Implementations
/// must not modify or dispose the stream, and must be deterministic — Parchment's byte-for-byte
/// output guarantee extends through them.
/// </para>
/// </remarks>
public interface IPageNumberResolver
{
    /// <summary>
    /// Maps bookmark name to its one-based page number. Bookmarks the resolver cannot place may be
    /// omitted; their fields keep the placeholder and stay marked for Word to update.
    /// </summary>
    Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel);
}
