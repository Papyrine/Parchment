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
/// <param name="written">
/// The markdown the render has produced so far, which is what says where a value lands. Fluid hands
/// the encoder a value and nothing else, so each render makes an encoder over its own output.
/// </param>
class MarkdownEncoder(StringBuilder? written) :
    TextEncoder
{
    /// <summary>
    /// For a value escaped away from any output, where there is no position to take into account.
    /// </summary>
    public static MarkdownEncoder Default { get; } = new(null);

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
        if (EndsADestination(value))
        {
            return EncodeDestination(value);
        }

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

    string EncodeSpan(CharSpan value)
    {
        if (EndsADestination(value))
        {
            return EncodeDestination(value);
        }

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

    // Only a value holding a space or a control character can end a destination early, and only
    // then is what has been written so far worth looking at.
    bool EndsADestination(CharSpan value) =>
        written != null &&
        value.IndexOfAnyInRange(char.MinValue, ' ') >= 0 &&
        InLinkDestination(written);

    static string EncodeDestination(CharSpan value)
    {
        var builder = new StringBuilder(value.Length + 16);
        WriteDestination(builder, value);
        return builder.ToString();
    }

    /// <summary>
    /// The longest unbroken run of source looked back over for the <c>](</c> that opens a
    /// destination. A url written in a template is far shorter; a line with no whitespace at all
    /// that is longer than this is not one.
    /// </summary>
    const int maxDestinationPrefix = 512;

    /// <summary>
    /// Whether the value about to be written lands in the destination of an inline link or image —
    /// <c>[text]({{ Url }})</c>, <c>![alt](images/{{ File }})</c> — going by the markdown already
    /// written.
    /// </summary>
    /// <remarks>
    /// Read from the output rather than worked out from the template source, so it holds whatever
    /// liquid produced the text before the value: a loop, a branch, an earlier value. Not inside a
    /// <c>capture</c>, which writes somewhere else. A destination
    /// the template wrote without whitespace is the only kind recognized — one in angle brackets
    /// that holds a space of its own is not, and does not need to be, since a space is allowed
    /// there.
    /// </remarks>
    static bool InLinkDestination(StringBuilder written)
    {
        var start = written.Length;
        var floor = Math.Max(0, start - maxDestinationPrefix);
        while (start > floor &&
               written[start - 1] > ' ')
        {
            start--;
        }

        if (start == floor &&
            floor > 0)
        {
            return false;
        }

        Span<char> buffer = stackalloc char[written.Length - start];
        written.CopyTo(start, buffer, buffer.Length);
        CharSpan prefix = buffer;

        var open = prefix.LastIndexOf("](");
        if (open < 0 ||
            EndsWithDanglingEscape(prefix[..open]))
        {
            return false;
        }

        // What is left is the destination so far. It is still open unless something in it closed
        // it: a > for the angle-bracket form, a ) with no ( of its own to match otherwise.
        var destination = prefix[(open + 2)..];
        var angled = destination.StartsWith('<');
        var depth = 0;
        for (var index = 0; index < destination.Length; index++)
        {
            switch (destination[index])
            {
                case '\\':
                    index++;
                    break;
                case '>' when angled:
                    return false;
                case '(' when !angled:
                    depth++;
                    break;
                case ')' when !angled:
                    if (depth == 0)
                    {
                        return false;
                    }

                    depth--;
                    break;
            }
        }

        return true;
    }

    /// <summary>
    /// <see cref="Write"/> for a value in a link destination, where a space or a control character
    /// ends the destination and the link with it: the raw markdown printed instead, which for an
    /// image put the file path into the document. Each is written as a character reference, which
    /// a destination resolves like the backslash escapes — so the url or path that reaches the
    /// link is the value exactly as bound.
    /// </summary>
    /// <remarks>
    /// That includes a line break. <see cref="Break"/> is html, and there is no html in a
    /// destination.
    /// </remarks>
    static void WriteDestination(StringBuilder builder, CharSpan value)
    {
        foreach (var character in value)
        {
            if (character <= ' ')
            {
                builder.Append("&#");
                builder.Append((int) character);
                builder.Append(';');
                continue;
            }

            if (escapable.Contains(character))
            {
                builder.Append('\\');
            }

            builder.Append(character);
        }
    }

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
