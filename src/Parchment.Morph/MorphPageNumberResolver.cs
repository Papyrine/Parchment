namespace Parchment;

/// <summary>
/// Resolves Parchment's page-dependent fields by laying the document out with Morph.
/// </summary>
/// <remarks>
/// Assign it to <c>TemplateStore.PageNumbers</c> and a rendered document arrives with its table of
/// contents built and its page references filled in, instead of leaving Word to compute them and
/// asking the reader for permission to do so on open:
/// <code>
/// var store = new TemplateStore
/// {
///     PageNumbers = new MorphPageNumberResolver()
/// };
/// </code>
/// <para>
/// The cost is a layout pass per render — pages are measured, not drawn, but the document is still
/// laid out in full. The other cost is fidelity: Morph is a very good approximation of Word's
/// layout, not Word, so a document it breaks differently gets page numbers that are confidently
/// wrong. Weigh that against the prompt it removes, and prefer it where the report's own
/// pagination is simple.
/// </para>
/// </remarks>
public class MorphPageNumberResolver(ImageExportOptions? options = null) :
    IPageNumberResolver
{
    static readonly SkiaDocumentConverter Converter = new();

    // Deterministic rendering by default: the same document must measure the same everywhere, or
    // Parchment's byte-for-byte output guarantee stops holding through this.
    readonly ImageExportOptions options = options ?? new()
    {
        DeterministicRendering = true
    };

    public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel) =>
        Task.FromResult(Converter.GetBookmarkPages(docx, options));
}
