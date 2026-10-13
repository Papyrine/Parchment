/// <summary>
/// The characters one child of a <c>w:r</c> stands for. Text is not only <c>w:t</c>: Word writes a
/// tab, a line break, a non-breaking hyphen and an inserted symbol as elements of their own, and
/// rewrites a tab character left inside <c>w:t</c> as <c>w:tab</c> the first time it saves the
/// document. Both editable readers (<see cref="EditableFieldReader"/> for plain text,
/// <see cref="OpenXmlHtmlSerializer"/> for rich text) read through here so neither drops one.
/// </summary>
static class RunCharacters
{
    /// <summary>
    /// Null for an element that carries no text. A line break of either spelling (<c>w:br</c>,
    /// <c>w:cr</c>) is a newline and a tab is a tab, which the caller may need to write as markup.
    /// </summary>
    public static string? For(OpenXmlElement element) =>
        element switch
        {
            Text text => text.Text,
            Break or CarriageReturn => "\n",
            TabChar => "\t",
            NoBreakHyphen => "‑",
            // w:softHyphen is deliberately nothing: it marks where a word may be hyphenated and
            // is not part of the word. As U+00AD it would be an invisible character in the value,
            // enough to stop a number parsing or two equal-looking strings comparing equal.
            SymbolChar symbol => Symbol(symbol),
            _ => null
        };

    // w:sym names a character by code point plus the font to draw it in. A value never carries a
    // font, so only the code point comes back. For a Unicode font that is the character itself.
    // For a symbol font (Symbol, Wingdings) Word stores the glyph in the private use area, F000
    // plus the font's own character code; that is returned as stored rather than guessed at — the
    // low byte is an index into the font, not the letter it happens to share a code with.
    static string? Symbol(SymbolChar symbol)
    {
        var code = symbol.Char?.Value;
        if (code == null ||
            !int.TryParse(code, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var codePoint) ||
            !Rune.IsValid(codePoint))
        {
            return null;
        }

        return XmlChars.Strip(char.ConvertFromUtf32(codePoint));
    }
}
