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

    // A solo substitution of a plain dotted path: what the docx flow requires of an Excelsior
    // token, expressed as the markdown equivalent. Anything more (filters, arithmetic) is not a
    // path the Excelsior getter can walk.
    static readonly Regex soloToken = new(
        @"^\{\{\s*(?<path>[A-Za-z_][A-Za-z0-9_]*(?:\s*\.\s*[A-Za-z_][A-Za-z0-9_]*)*)\s*\}\}$",
        RegexOptions.Compiled);

    // Any substitution, to find the ones that are NOT alone on their line.
    static readonly Regex anyToken = new(@"\{\{(?<body>[^{}]*)\}\}", RegexOptions.Compiled);

    /// <summary>
    /// Replaces each Excelsior token with a marker, returning the rewritten source and what each
    /// marker stands for. Returns the source unchanged when the model declares no tables.
    /// </summary>
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
        var lines = markdown.Split('\n');
        for (var index = 0; index < lines.Length; index++)
        {
            // Split on \n alone so a CRLF source keeps its \r; the trim below takes it off the
            // comparison and the line is replaced wholesale when it matches.
            var line = lines[index];
            var trimmed = line.Trim();

            var match = soloToken.Match(trimmed);
            if (match.Success)
            {
                var path = Normalize(match.Groups["path"].Value);
                if (tables.TryGet(path, out _))
                {
                    var marker = MarkerFor(placeholders.Count);
                    placeholders.Add(new(marker, path));
                    lines[index] = marker;
                    continue;
                }
            }

            GuardUnrenderableToken(trimmed, tables, templateName);
        }

        if (placeholders.Count == 0)
        {
            return (markdown, []);
        }

        return (string.Join("\n", lines), placeholders);
    }

    /// <summary>
    /// Rejects an Excelsior token the rewrite above could not take.
    /// </summary>
    /// <remarks>
    /// Left to render, such a token falls through to plain Fluid output and writes the collection's
    /// ToString into the document — a wrong document rather than a failed one. The docx flow rejects
    /// both shapes too, as PARCH007 and PARCH008, so this keeps the two flows answering the same way.
    /// </remarks>
    static void GuardUnrenderableToken(string trimmedLine, ExcelsiorTableMap tables, string templateName)
    {
        foreach (Match match in anyToken.Matches(trimmedLine))
        {
            var body = match.Groups["body"].Value.Trim();

            // The leading path, up to whatever ends it — a filter pipe, an operator, a space.
            var end = body.IndexOfAny([' ', '|', '=', '<', '>', '!', '+', '-', '*', '/', '[', '(']);
            var leading = Normalize(end < 0 ? body : body[..end]);
            if (leading.Length == 0 ||
                !tables.TryGet(leading, out _))
            {
                continue;
            }

            // Alone on the line, so what disqualified it was the expression rather than its
            // company: the getter walks the model directly and never sees a filter's output.
            if (match.Value == trimmedLine)
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"'{trimmedLine}' renders the [ExcelsiorTable] '{leading}' as a Word table, which " +
                    "is built from the model rather than from Fluid's output, so the token has to be " +
                    "a plain member access. Filters and expressions cannot be applied to it.");
            }

            throw new ParchmentRegistrationException(
                templateName,
                $"'{{{{ {leading} }}}}' renders an [ExcelsiorTable] as a Word table, which replaces " +
                "the whole block it sits in, so it has to be alone on its line. It currently shares " +
                $"a line with other content: {trimmedLine}");
        }
    }

    // Fluid tolerates spaces around the dots; the map is keyed on the bare path.
    static string Normalize(string path) =>
        path.Replace(" ", "").Replace("\t", "");

    /// <summary>
    /// Swaps each marker's paragraph for the table it stands for.
    /// </summary>
    public static void Apply(
        Body body,
        IReadOnlyList<Placeholder> placeholders,
        ExcelsiorTableMap tables,
        object model,
        MainDocumentPart mainPart,
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

                Replace(host, data, entry, mainPart);
            }
        }
    }

    static void Replace(Paragraph host, object? data, ExcelsiorTableEntry entry, MainDocumentPart mainPart)
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

        parent.InsertBefore(table, host);

        // Word merges adjacent tables and mishandles a table that ends a body, so the host
        // paragraph is emptied and kept as the separator rather than removed.
        host.RemoveAllChildren();
    }
}
