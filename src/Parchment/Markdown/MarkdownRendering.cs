/// <summary>
/// Entry point for rendering a markdown string into a list of OpenXML elements suitable for
/// splicing into a Word document body, header, footer, or other content host.
/// </summary>
static class MarkdownRendering
{
    public static IReadOnlyList<OpenXmlElement> Render(string markdown, MainDocumentPart mainPart, WordNumberingState numbering, ImagePolicies imagePolicies, int headingOffset)
    {
        var document = Markdown.Parse(markdown, MarkdigPipeline.Pipeline);
        var renderer = new OpenXmlMarkdownRenderer(mainPart, numbering, imagePolicies, headingOffset);
        MarkAlignedPipeTables(markdown, document, renderer);
        renderer.Render(document);
        return renderer.Drain();
    }

    static void MarkAlignedPipeTables(string markdown, MarkdownDocument document, OpenXmlMarkdownRenderer renderer)
    {
        string[]? lines = null;
        foreach (var table in document.Descendants<Markdig.Extensions.Tables.Table>())
        {
            if (table.ColumnDefinitions.Count == 0)
            {
                continue;
            }

            float totalPct = 0;
            foreach (var column in table.ColumnDefinitions)
            {
                totalPct += column.Width;
            }

            // No width hints from Markdig — no decision to make.
            if (totalPct <= 0)
            {
                continue;
            }

            lines ??= SplitLines(markdown);
            var headerLine = table.Line;
            // Separator row was consumed during parsing; bodyCount is the rows remaining (header
            // included as the first AST row, body rows after it). The original source separator
            // sits at headerLine + 1, body rows at headerLine + 2 ... headerLine + 1 + bodyCount.
            var bodyCount = table.OfType<Markdig.Extensions.Tables.TableRow>().Count() - 1;
            var lastBodyLine = headerLine + 1 + bodyCount;
            if (headerLine < 0 || lastBodyLine >= lines.Length)
            {
                continue;
            }

            // Only pipe tables (separator-based) get this heuristic. Grid tables (+---+ borders)
            // express layout intent intrinsically and should always honour their inferred widths.
            var trimmedHeader = lines[headerLine].TrimStart();
            if (trimmedHeader.Length == 0 || trimmedHeader[0] != '|')
            {
                continue;
            }

            var referencePipes = ExtractPipePositions(lines[headerLine]);
            if (referencePipes.Count == 0)
            {
                continue;
            }

            var aligned = true;
            for (var i = headerLine + 1; i <= lastBodyLine; i++)
            {
                if (!PipePositionsMatch(referencePipes, ExtractPipePositions(lines[i])))
                {
                    aligned = false;
                    break;
                }
            }

            if (aligned)
            {
                renderer.SkipColumnWidths.Add(table);
            }
        }
    }

    static string[] SplitLines(string source)
    {
        var raw = source.Split('\n');
        for (var i = 0; i < raw.Length; i++)
        {
            if (raw[i].Length > 0 && raw[i][^1] == '\r')
            {
                raw[i] = raw[i][..^1];
            }
        }

        return raw;
    }

    static List<int> ExtractPipePositions(string line)
    {
        var result = new List<int>();
        var escaped = false;
        for (var i = 0; i < line.Length; i++)
        {
            if (escaped)
            {
                escaped = false;
                continue;
            }

            if (line[i] == '\\')
            {
                escaped = true;
                continue;
            }

            if (line[i] == '|')
            {
                result.Add(i);
            }
        }

        return result;
    }

    static bool PipePositionsMatch(List<int> a, List<int> b)
    {
        if (a.Count != b.Count)
        {
            return false;
        }

        for (var i = 0; i < a.Count; i++)
        {
            if (a[i] != b[i])
            {
                return false;
            }
        }

        return true;
    }
}
