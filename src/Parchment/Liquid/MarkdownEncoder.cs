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
    public const string Break = "<br />";

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

    /// <summary>
    /// What the first character of a value is written as when no backslash can make it literal:
    /// a character reference, which CommonMark reads as the character and as nothing else.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A space or tab, because indentation is block syntax: four columns of it make the line an
    /// indented code block, and less than that is stripped. A reference cannot stand in for
    /// structural whitespace, so the value keeps its indent and the line gains none. Only the
    /// first character needs it — once the line opens with something that is not whitespace, the
    /// rest of the run is ordinary text.
    /// </para>
    /// <para>
    /// An opening bracket, because <c>[TOC]</c> and <c>[SECTION]</c> are not syntax at all: they
    /// are plain text that <see cref="TableOfContents"/> and <see cref="SectionBreaks"/> recognise
    /// after parsing, and an escaped bracket is still the text of a bracket. Both ask for a
    /// paragraph of nothing but literals, and a reference is parsed into an inline of its own, so
    /// a paragraph a value opens is never a marker. A bracket further in cannot open one.
    /// </para>
    /// <para>
    /// A later line of the value needs neither: its newline became <see cref="Break"/>, so it
    /// never starts a line of the source, and the paragraph it is in already holds an inline that
    /// is not a literal.
    /// </para>
    /// </remarks>
    static string? LeadingReference(char character) =>
        character switch
        {
            ' ' => "&#32;",
            '\t' => "&#9;",
            '[' => "&#91;",
            _ => null
        };

    public override int MaxOutputCharactersPerInputCharacter => Break.Length;

    public override bool WillEncode(int unicodeScalar) =>
        unicodeScalar is '\r' or '\n' ||
        (unicodeScalar <= char.MaxValue && escapable.Contains((char) unicodeScalar));

    public override string Encode(string value)
    {
        if (IsUntouched(value))
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
        if (IsUntouched(value))
        {
            return value.ToString();
        }

        var builder = new StringBuilder(value.Length + 16);
        Write(builder, value);
        return builder.ToString();
    }

    static bool IsUntouched(CharSpan value) =>
        value.IndexOfAny(encodable) < 0 &&
        !StartsWithIndent(value);

    // A bracket is in encodable already. Indentation is the one thing that needs work and holds
    // no character from that set.
    static bool StartsWithIndent(CharSpan value) =>
        value.Length > 0 &&
        value[0] is ' ' or '\t';

    static void Write(StringBuilder builder, CharSpan value)
    {
        var start = 0;
        if (value.Length > 0 &&
            LeadingReference(value[0]) is { } reference)
        {
            builder.Append(reference);
            start = 1;
        }

        for (var index = start; index < value.Length; index++)
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

    /// <summary>
    /// Undoes <see cref="Encode(string)"/>'s backslashes for the one construct that reads them
    /// verbatim and still has to work: an autolink, whose text is also its target. A backslash that
    /// escapes nothing is kept.
    /// </summary>
    public static string Unescape(CharSpan value)
    {
        if (value.IndexOf('\\') < 0)
        {
            return value.ToString();
        }

        var builder = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] == '\\' &&
                index + 1 < value.Length &&
                escapable.Contains(value[index + 1]))
            {
                index++;
            }

            builder.Append(value[index]);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Whether the last character is a backslash left over once every escape before it has been
    /// paired off — one that escapes whatever followed the span, not anything inside it.
    /// </summary>
    public static bool EndsWithDanglingEscape(CharSpan value)
    {
        var count = 0;
        while (count < value.Length &&
               value[^(count + 1)] == '\\')
        {
            count++;
        }

        return count % 2 == 1;
    }

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
