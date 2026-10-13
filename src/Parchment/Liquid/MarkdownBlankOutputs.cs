/// <summary>
/// Removes the emphasis a markdown template wrapped around a token whose value turned out blank, so
/// <c>**{{ Name }}**</c> prints nothing for an empty <c>Name</c> rather than <c>****</c>.
/// </summary>
/// <remarks>
/// <para>
/// The template is assembled as text and only then parsed, and markdown has no empty emphasis: a
/// delimiter run with nothing in it is literal text, and <c>****</c> or <c>___</c> alone on a line
/// is a thematic break. Nothing downstream can tell those delimiters from ones the author meant to
/// print, so the place a blank value was written is recorded while liquid renders and the
/// delimiters around exactly that place are dropped before Markdig sees the text. What is left is
/// what the same line gives with the token bare.
/// </para>
/// <para>
/// Positions rather than a placeholder character: nothing is added to the text, so nothing has to
/// be found and removed again, and a value cannot forge one. Only markdown source assembled from a
/// template goes through here — a <c>[Markdown]</c> value or a <c>markdown</c> filter in the docx
/// flow is source its author wrote, and a literal <c>****</c> there is theirs.
/// </para>
/// </remarks>
static class MarkdownBlankOutputs
{
    const string ambientKey = "Parchment.MarkdownBlankOutputs";

    /// <summary>
    /// Where the blank values of one render landed in the text <paramref name="writer"/> collects.
    /// </summary>
    public sealed class Tracker(StringWriter writer)
    {
        public StringWriter Writer { get; } = writer;

        // In the order written, so ascending by start.
        public List<(int Start, int Length)> Spans { get; } = [];
    }

    /// <summary>
    /// Wraps every <c>{{ }}</c> of <paramref name="template"/> so a render can tell which ones
    /// wrote nothing but whitespace.
    /// </summary>
    public static IFluidTemplate Instrument(IFluidTemplate template) =>
        new Rewriter().VisitTemplate(template);

    public static Tracker Begin(TemplateContext context, StringWriter writer)
    {
        var tracker = new Tracker(writer);
        context.AmbientValues[ambientKey] = tracker;
        return tracker;
    }

    sealed class Rewriter :
        AstRewriter
    {
        protected override Statement VisitOutputStatement(OutputStatement outputStatement) =>
            new TrackedOutputStatement(outputStatement);
    }

    /// <summary>
    /// Measures what the wrapped statement wrote rather than evaluating its expression a second
    /// time: a filter can park a value as a side effect, and what counts is the text that reached
    /// the source, whatever type produced it.
    /// </summary>
    sealed class TrackedOutputStatement(OutputStatement inner) :
        Statement
    {
        public override async ValueTask<Completion> WriteToAsync(TextWriter writer, TextEncoder encoder, TemplateContext context)
        {
            // A capture block renders into a writer of its own, and where its text ends up is
            // decided later, so only what goes straight to the render's writer is tracked.
            if (!context.AmbientValues.TryGetValue(ambientKey, out var value) ||
                value is not Tracker tracker ||
                !ReferenceEquals(tracker.Writer, writer))
            {
                return await inner.WriteToAsync(writer, encoder, context).ConfigureAwait(false);
            }

            var builder = tracker.Writer.GetStringBuilder();
            var start = builder.Length;
            var completion = await inner.WriteToAsync(writer, encoder, context).ConfigureAwait(false);

            if (IsBlank(builder, start))
            {
                tracker.Spans.Add((start, builder.Length - start));
            }

            return completion;
        }

        // MarkdownEncoder writes the first space or tab of a value as a character reference, so
        // a blank value opens with one.
        static bool IsBlank(StringBuilder builder, int start)
        {
            for (var index = start + LeadingSpaceReference(builder, start); index < builder.Length; index++)
            {
                if (!IsInlineSpace(builder[index]))
                {
                    return false;
                }
            }

            return true;
        }

        static int LeadingSpaceReference(StringBuilder builder, int start)
        {
            foreach (var reference in spaceReferences)
            {
                if (builder.Length - start < reference.Length)
                {
                    continue;
                }

                var matches = true;
                for (var index = 0; index < reference.Length; index++)
                {
                    if (builder[start + index] != reference[index])
                    {
                        matches = false;
                        break;
                    }
                }

                if (matches)
                {
                    return reference.Length;
                }
            }

            return 0;
        }

        static readonly string[] spaceReferences = ["&#32;", "&#9;"];
    }

    /// <summary>
    /// Returns <paramref name="tracker"/>'s text with the emphasis around each blank value removed.
    /// </summary>
    public static string Strip(Tracker tracker)
    {
        var builder = tracker.Writer.GetStringBuilder();
        var spans = tracker.Spans;
        if (spans.Count == 0)
        {
            return builder.ToString();
        }

        // Most blank values sit nowhere near a delimiter, so the common case is a read-only pass
        // over the string with no second copy made.
        var text = builder.ToString();
        StringBuilder? edited = null;

        // Last to first, so a removal never moves a span still to be visited. limit is where the
        // text removed so far begins: a span reaching past it had its tail removed with it.
        var limit = text.Length;
        var index = spans.Count - 1;
        while (index >= 0)
        {
            var (start, length) = spans[index];
            var end = Math.Min(start + length, limit);
            index--;

            // `**{{ First }} {{ Last }}**` with both blank is one blank stretch, not two values
            // each with a delimiter on one side only.
            while (index >= 0 &&
                   IsBlankBetween(edited, text, spans[index].Start + spans[index].Length, start))
            {
                start = spans[index].Start;
                index--;
            }

            if (start > end)
            {
                continue;
            }

            var (from, to) = Widen(edited, text, start, end);
            if (from == start)
            {
                continue;
            }

            edited ??= new(text);
            edited.Remove(from, to - from);
            limit = from;
        }

        if (edited == null)
        {
            return text;
        }

        return edited.ToString();
    }

    /// <summary>
    /// Grows a blank stretch outwards over each pair of delimiter runs that encloses it, innermost
    /// first, so <c>~~**{{ A }}**~~</c> goes entirely.
    /// </summary>
    static (int From, int To) Widen(StringBuilder? edited, string text, int start, int end)
    {
        var length = edited?.Length ?? text.Length;
        while (start > 0 && end < length)
        {
            var delimiter = At(edited, text, start - 1);
            if (delimiter is not ('*' or '_' or '~' or '`') ||
                At(edited, text, end) != delimiter)
            {
                break;
            }

            var opening = 1;
            while (start - opening > 0 &&
                   At(edited, text, start - opening - 1) == delimiter)
            {
                opening++;
            }

            var closing = 1;
            while (end + closing < length &&
                   At(edited, text, end + closing) == delimiter)
            {
                closing++;
            }

            // Only a run that could open emphasis and one that could close it. `**a**{{ B }}*c*`
            // has a closing run before the token and an opening one after it, and both belong to
            // their neighbours. A backslash before the run means the template escaped it.
            var before = start - opening - 1;
            var after = end + closing;
            if ((before >= 0 && !IsBoundary(At(edited, text, before), leading: true)) ||
                (after < length && !IsBoundary(At(edited, text, after), leading: false)))
            {
                break;
            }

            // A code span closes only on a run of the same length as the one that opened it.
            var count = Math.Min(opening, closing);
            if (delimiter == '`' &&
                opening != closing)
            {
                break;
            }

            start -= count;
            end += count;

            if (opening != closing)
            {
                break;
            }
        }

        return (start, end);
    }

    static bool IsBoundary(char character, bool leading)
    {
        if (leading && character == '\\')
        {
            return false;
        }

        return char.IsWhiteSpace(character) ||
               char.IsPunctuation(character) ||
               char.IsSymbol(character);
    }

    static bool IsBlankBetween(StringBuilder? edited, string text, int from, int to)
    {
        for (var index = from; index < to; index++)
        {
            if (!IsInlineSpace(At(edited, text, index)))
            {
                return false;
            }
        }

        return true;
    }

    // A line break is not blank here: it ends the line the delimiters are on, so nothing either
    // side of it can pair up.
    static bool IsInlineSpace(char character) =>
        character is ' ' or '\t';

    static char At(StringBuilder? edited, string text, int index)
    {
        if (edited != null)
        {
            return edited[index];
        }

        return text[index];
    }
}
