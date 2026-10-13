namespace Parchment;

/// <summary>
/// Narrow context passed to <see cref="OpenXmlToken"/> render delegates. Kept intentionally
/// small so the public API does not lock users to a specific DocumentFormat.OpenXml version.
/// </summary>
public interface IOpenXmlContext
{
    /// <summary>
    /// The main document part, whichever part the token sits in. A relationship id resolves
    /// against the part holding the element that carries it, so when the token sits in a header,
    /// footer, footnote or endnote, a relationship added here that the token's content refers to
    /// is moved to that part once the delegate returns, and the content pointed at its new id.
    /// </summary>
    MainDocumentPart MainPart { get; }
    int CurrentHeadingLevel { get; }

    /// <summary>
    /// The paragraph being replaced by the token's output, when the token sits as a solo
    /// substitution that consumes the whole host paragraph. Null for inline / non-structural
    /// renders. Tokens can read the host's <c>pStyle</c> (or other properties) to make the
    /// produced content inherit the surrounding context — e.g. so a list rendered inside a
    /// table cell picks up the cell's font instead of falling back to the global default.
    /// </summary>
    Paragraph? HostParagraph { get; }

    /// <summary>
    /// Adds an image to the part the token sits in — the main document part, or the header,
    /// footer, footnotes or endnotes part — and returns the relationship id to refer to it by.
    /// </summary>
    string AddImagePart(byte[] bytes, string contentType);
    bool TryGetStyle(string styleId, out StyleType styleType);
    int CreateBulletNumbering();
    int CreateOrderedNumbering(NumberFormatValues numberFormat);
}
