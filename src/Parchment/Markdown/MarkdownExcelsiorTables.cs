/// <summary>
/// Renders <c>[ExcelsiorTable]</c> members as Word tables in the markdown flow.
/// </summary>
/// <remarks>
/// The docx flow swaps the token's host paragraph for the built table. A markdown template has no
/// host paragraph at substitution time — liquid renders to markdown source and Markdig parses that
/// source afterwards — so the swap happens on the other side of the parse instead: the token is
/// replaced with a marker before the source is parsed, and the paragraph the marker lands in is
/// replaced with the table once the markdown has become OpenXML.
///
/// Rewriting at registration rather than at render keeps it off the hot path: the marked-up source
/// is parsed into the cached <c>IFluidTemplate</c>, so a render sees markers and nothing else.
/// </remarks>
static class MarkdownExcelsiorTables
{
    /// <summary>
    /// A rewritten token, and the model path whose table takes its place.
    /// </summary>
    public sealed record Placeholder(string Marker, string DottedPath);

    /// <summary>
    /// Private-use characters, so no rendered value can collide with a marker. A model value that
    /// happened to equal the marker text would otherwise be swapped for a table.
    /// </summary>
    static string MarkerFor(int index) =>
        $"parchment-table-{index}";

    /// <summary>
    /// Replaces each Excelsior token with a marker, returning the rewritten source and what each
    /// marker stands for. Returns the source unchanged when the model declares no tables.
    /// </summary>
    /// <remarks>
    /// A token is taken when it is a plain path alone on its line: what the docx flow requires of
    /// an Excelsior token, expressed as the markdown equivalent. Both are read off the token as
    /// Fluid parses it (see <see cref="MarkdownOutputTokens"/>), so every spelling of it counts.
    /// </remarks>
    public static (string Markdown, IReadOnlyList<Placeholder> Placeholders) Rewrite(
        string markdown,
        ExcelsiorTableMap tables,
        string templateName)
    {
        if (tables.IsEmpty)
        {
            return (markdown, []);
        }

        var placeholders = new List<Placeholder>();
        var loops = MarkdownLoopScope.Scan(markdown);
        var builder = new StringBuilder();
        var copied = 0;
        foreach (var token in MarkdownOutputTokens.Scan(markdown, SharedFluid.Parser))
        {
            // A token that starts with a loop variable reads the loop item, whatever root path
            // it happens to spell: Fluid renders it, and neither rule below is its to break.
            if (token.Path == null ||
                loops.Shadows(token.Start, token.Path) ||
                !tables.TryGet(token.Path, out _))
            {
                continue;
            }

            var alone = MarkdownOutputTokens.IsAlone(markdown, token, out var lineStart, out var lineEnd);
            GuardRenderable(markdown, token, alone, lineStart, lineEnd, templateName);

            var marker = MarkerFor(placeholders.Count);
            placeholders.Add(new(marker, token.Path));

            // The whole line goes, not only the token: indentation left in front of the marker
            // could make it a code block, and a CRLF source's \r has nothing left to end.
            builder.Append(markdown, copied, lineStart - copied);
            AppendMarker(builder, marker, token);
            copied = lineEnd;
        }

        if (placeholders.Count == 0)
        {
            return (markdown, []);
        }

        builder.Append(markdown, copied, markdown.Length - copied);
        return (builder.ToString(), placeholders);
    }

    /// <summary>
    /// Writes the marker in place of its token. A token with whitespace control becomes an output
    /// of the marker carrying the same dashes, so it still trims what it trimmed; <c>raw</c> keeps
    /// <c>MarkdownEncoder</c> off the marker's punctuation.
    /// </summary>
    static void AppendMarker(StringBuilder builder, string marker, MarkdownOutputToken token)
    {
        if (!token.TrimLeft &&
            !token.TrimRight)
        {
            builder.Append(marker);
            return;
        }

        builder.Append(token.TrimLeft ? "{{- '" : "{{ '");
        builder.Append(marker);
        builder.Append(token.TrimRight ? "' | raw -}}" : "' | raw }}");
    }

    /// <summary>
    /// Rejects an Excelsior token the rewrite above cannot take.
    /// </summary>
    /// <remarks>
    /// Left to render, such a token falls through to plain Fluid output and writes the collection's
    /// ToString into the document — a wrong document rather than a failed one. The docx flow rejects
    /// both shapes too, as PARCH007 and PARCH008, so this keeps the two flows answering the same way.
    /// </remarks>
    static void GuardRenderable(
        string markdown,
        MarkdownOutputToken token,
        bool alone,
        int lineStart,
        int lineEnd,
        string templateName)
    {
        if (!alone)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"'{{{{ {token.Path} }}}}' renders an [ExcelsiorTable] as a Word table, which replaces " +
                "the whole block it sits in, so it has to be alone on its line. It currently shares " +
                $"a line with other content: {markdown[lineStart..lineEnd].Trim()}");
        }

        // Alone on the line, so what disqualifies it is the expression rather than its company:
        // the getter walks the model directly and never sees a filter's output.
        if (!token.IsPlain)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"'{markdown[token.Start..token.End]}' renders the [ExcelsiorTable] '{token.Path}' as a Word table, which " +
                "is built from the model rather than from Fluid's output, so the token has to be " +
                "a plain member access. Filters and expressions cannot be applied to it.");
        }
    }
    /// <summary>
    /// Swaps each marker's paragraph for the table it stands for.
    /// </summary>
    public static void Apply(
        Body body,
        IReadOnlyList<Placeholder> placeholders,
        ExcelsiorTableMap tables,
        object model,
        MainDocumentPart mainPart,
        BookmarkIds bookmarks,
        string templateName)
    {
        if (placeholders.Count == 0)
        {
            return;
        }

        foreach (var placeholder in placeholders)
        {
            if (!tables.TryGet(placeholder.DottedPath, out var entry))
            {
                continue;
            }

            var data = entry.Getter(model);

            // Materialized before mutating: the swap edits the tree being walked. A marker inside a
            // loop appears once per iteration, and a marker inside a false conditional not at all,
            // so neither the count nor its being zero is an error.
            var hosts = body.Descendants<Paragraph>()
                .Where(_ => _.InnerText.Contains(placeholder.Marker, StringComparison.Ordinal))
                .ToList();

            foreach (var host in hosts)
            {
                if (host.InnerText.Trim() != placeholder.Marker)
                {
                    throw new ParchmentRenderException(
                        templateName,
                        $"'{{{{ {placeholder.DottedPath} }}}}' has to be alone in its block: the " +
                        "table replaces the block, so the text sharing it would be discarded. " +
                        $"Markdown parsed it into a paragraph reading: {host.InnerText.Trim()}");
                }

                Replace(host, data, entry, mainPart, bookmarks);
            }
        }
    }

    static void Replace(Paragraph host, object? data, ExcelsiorTableEntry entry, MainDocumentPart mainPart, BookmarkIds bookmarks)
    {
        var parent = host.Parent!;

        if (data == null)
        {
            // Nothing to tabulate. The marker still has to go, or it renders as text.
            host.Remove();
            return;
        }

        var table = ExcelsiorTableBridge.BuildTable(
            entry.ElementType,
            data,
            mainPart,
            entry.HeadingParagraphStyle,
            entry.BodyParagraphStyle,
            entry.TableStyle,
            entry.Configure);

        // Like everything else that joins the body. Excelsior writes no bookmarks of its own, but a
        // Configure hook is caller code, and the table of contents numbers its entries from the
        // same counter on the strength of nothing in the body having bypassed it.
        bookmarks.Renumber(table);
        parent.InsertBefore(table, host);

        // Word merges adjacent tables and mishandles a table that ends a body, so the host
        // paragraph is emptied and kept as the separator rather than removed.
        host.RemoveAllChildren();
    }
}
