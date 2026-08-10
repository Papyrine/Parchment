using System.Buffers;
using System.Text.Encodings.Web;

namespace Parchment;

/// <summary>
/// The encoder the markdown flow renders liquid through, so a bound value lands in the document as
/// the text it is rather than as markdown source.
/// </summary>
/// <remarks>
/// <para>
/// Fluid writes every <c>{{ }}</c> result through the encoder handed to <c>RenderAsync</c>. The
/// markdown flow used to hand it <c>NullEncoder</c>, so a value carrying a <c>|</c> broke the table
/// row it sat in, and one carrying <c>{.Heading1}</c> restyled the paragraph around it. Escaping is
/// the default for the same reason it is in html templating: the template author knows which sites
/// are markup, the model does not, and the failure is silent.
/// </para>
/// <para>
/// Only output is encoded. Comparisons, filter inputs and loop sources see the raw value, so
/// <c>{% if Title == "A.B" %}</c> still matches what the model holds.
/// </para>
/// <para>
/// A site opts out with <c>| raw</c>, with the <c>markdown</c> filter, or by typing the member as
/// <see cref="TokenValue"/> and returning <see cref="MarkdownToken"/> or <see cref="HtmlToken"/> —
/// the type saying "this value is markup" is what turns escaping off. See <c>Filters</c> for the
/// ones that opt themselves out.
/// </para>
/// </remarks>
class MarkdownEncoder :
    TextEncoder
{
    public static MarkdownEncoder Default { get; } = new();

    /// <summary>
    /// A newline becomes inline html rather than being escaped or folded to a space: a line break
    /// in bound content is a line break in the document, and <c>HtmlInlineRenderer</c> already
    /// turns <c>&lt;br&gt;</c> into a <c>w:br</c>.
    /// </summary>
    /// <remarks>
    /// Inline html is the only form that survives everywhere a value can land. A hard break (two
    /// trailing spaces, or a trailing backslash) needs a real newline in the source, which ends the
    /// row it is in when the value sits in a table cell — the original breakage. A blank line would
    /// end the host paragraph, taking the value out of its style and orphaning any <c>{.Style}</c>
    /// attribute attached to it.
    /// </remarks>
    const string Break = "<br />";

    /// <summary>
    /// Characters a backslash makes literal. Every one is ASCII punctuation, which CommonMark
    /// always treats as an escape — before anything else the backslash stays a backslash and would
    /// print, so this set must not grow past that.
    /// </summary>
    /// <remarks>
    /// Covers what <see cref="MarkdigPipeline"/> can read as syntax: pipe tables (<c>|</c>), generic
    /// attributes (<c>{}</c>), emphasis extras (<c>^</c>, <c>~</c>, <c>=</c>), setext underlines
    /// (<c>=</c>, <c>-</c>) and the CommonMark core. Quotes and dashes are deliberately absent —
    /// smarty pants rewriting them is a typographic nicety, not a structural hazard.
    /// </remarks>
    static readonly SearchValues<char> escapable = SearchValues.Create(@"\`*_~^[]()<>&#|{}!+-.=");

    /// <summary>
    /// <see cref="escapable"/> plus the line-break characters, so one vectorized scan decides
    /// whether a value needs any work at all. Most do not.
    /// </summary>
    static readonly SearchValues<char> encodable = SearchValues.Create("\r\n" + @"\`*_~^[]()<>&#|{}!+-.=");

    public override int MaxOutputCharactersPerInputCharacter => Break.Length;

    public override bool WillEncode(int unicodeScalar) =>
        unicodeScalar is '\r' or '\n' ||
        (unicodeScalar <= char.MaxValue && escapable.Contains((char) unicodeScalar));

    public override string Encode(string value)
    {
        if (value.AsSpan().IndexOfAny(encodable) < 0)
        {
            return value;
        }

        var builder = new StringBuilder(value.Length + 16);
        Write(builder, value);
        return builder.ToString();
    }

    public override void Encode(TextWriter output, string value, int startIndex, int characterCount) =>
        output.Write(EncodeSpan(value.AsSpan(startIndex, characterCount)));

    public override void Encode(TextWriter output, char[] value, int startIndex, int characterCount) =>
        output.Write(EncodeSpan(value.AsSpan(startIndex, characterCount)));

    static string EncodeSpan(CharSpan value)
    {
        if (value.IndexOfAny(encodable) < 0)
        {
            return value.ToString();
        }

        var builder = new StringBuilder(value.Length + 16);
        Write(builder, value);
        return builder.ToString();
    }

    static void Write(StringBuilder builder, CharSpan value)
    {
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];

            if (character is '\r' or '\n')
            {
                // A CRLF is one line break. Swallow the LF so the pair does not write two.
                if (character == '\r' &&
                    index + 1 < value.Length &&
                    value[index + 1] == '\n')
                {
                    index++;
                }

                builder.Append(Break);
                continue;
            }

            if (escapable.Contains(character))
            {
                builder.Append('\\');
            }

            builder.Append(character);
        }
    }

    /// <summary>
    /// Escapes a value that is being assembled into markdown source by a filter rather than written
    /// straight to the output — the filter's own markers have to stay syntax, so it escapes the
    /// parts that came from the model itself and returns a value the encoder leaves alone.
    /// </summary>
    public static string EscapeValue(string value) =>
        Default.Encode(value);

    // Reached only through base-class entry points this type does not override; every one Fluid
    // uses is overridden above. Kept correct for a single character, which is all the signature can
    // express — a scalar at a time cannot see a CRLF pair, so this writes a break for each half.
    public override unsafe int FindFirstCharacterToEncode(char* text, int textLength) =>
        new CharSpan(text, textLength).IndexOfAny(encodable);

    public override unsafe bool TryEncodeUnicodeScalar(int unicodeScalar, char* buffer, int bufferLength, out int charsWritten)
    {
        charsWritten = 0;

        if (unicodeScalar is '\r' or '\n')
        {
            if (bufferLength < Break.Length)
            {
                return false;
            }

            Break.CopyTo(new(buffer, bufferLength));
            charsWritten = Break.Length;
            return true;
        }

        if (bufferLength < 2)
        {
            return false;
        }

        buffer[0] = '\\';
        buffer[1] = (char) unicodeScalar;
        charsWritten = 2;
        return true;
    }
}
