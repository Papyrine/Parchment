/// <summary>
/// Replaces a token's character range inside a live paragraph with a sequence of inline elements
/// (the editable-field permStart / sdt / permEnd triple), preserving every sibling in place.
///
/// <see cref="ParagraphSplicer"/>'s clone-and-trim approach is unsuitable here: it distributes
/// zero-width elements (bookmarks, perm-range markers, emptied sdt shells) into <em>both</em>
/// halves, which corrupts previously-spliced editable fields when a paragraph hosts more than
/// one. This splicer instead splits runs at the exact token boundaries and inserts between them,
/// touching nothing outside the token's character range.
///
/// Tokens are processed in reverse-offset order (see <c>ScopeTreeRunner</c>), so any previously
/// inserted field sits at a higher offset than the current token and is never walked into.
/// </summary>
static class EditableSplicer
{
    public static void Insert(
        Paragraph host,
        int offset,
        int length,
        Func<RunProperties?, IReadOnlyList<OpenXmlElement>> factory)
    {
        // Formatting rule: the run owning the first character of the token wins — same as
        // ParagraphText.Replace.
        var firstRun = RunAt(host, offset);
        var sitePr = (RunProperties?)firstRun?.RunProperties?.CloneNode(true);
        var produced = factory(sitePr);

        var insertAfter = EnsureBoundary(host, offset);
        EnsureBoundary(host, offset + length);
        RemoveCoveredTexts(host, offset, offset + length);

        var cursor = insertAfter ?? host.ParagraphProperties;
        foreach (var element in produced)
        {
            cursor = cursor == null
                ? host.InsertAt(element, 0)
                : host.InsertAfter(element, cursor);
        }
    }

    static Run? RunAt(Paragraph host, int offset)
    {
        var consumed = 0;
        foreach (var text in host.Descendants<Text>())
        {
            var end = consumed + text.Text.Length;
            if (offset >= consumed && offset < end)
            {
                return text.Ancestors<Run>().FirstOrDefault();
            }

            consumed = end;
        }

        return null;
    }

    /// <summary>
    /// Ensures a paragraph-level element boundary exists at <paramref name="absOffset"/> and
    /// returns the direct child of <paramref name="host"/> after which content at that offset
    /// begins (null when the offset is at the very start of the paragraph's content).
    /// </summary>
    static OpenXmlElement? EnsureBoundary(Paragraph host, int absOffset)
    {
        var consumed = 0;
        OpenXmlElement? lastTopLevelBefore = null;
        foreach (var text in host.Descendants<Text>().ToList())
        {
            var start = consumed;
            var end = consumed + text.Text.Length;
            consumed = end;

            if (end <= absOffset)
            {
                lastTopLevelBefore = TopLevel(host, text);
                continue;
            }

            var run = (Run)text.Parent!;
            if (start >= absOffset)
            {
                // Boundary falls exactly before this text. If the run holds earlier content,
                // split the run so the boundary is at paragraph level.
                if (HasContentBefore(run, text))
                {
                    SplitRunBefore(run, text);
                    return TopLevel(host, run);
                }

                return TopLevel(host, run).PreviousSibling();
            }

            // Boundary falls inside this text — split the text, then the run.
            var local = absOffset - start;
            var value = text.Text;
            var tail = new Text(value[local..]);
            text.Text = value[..local];
            if (text.Text.Length > 0)
            {
                text.Space = SpaceProcessingModeValues.Preserve;
            }

            if (tail.Text.Length > 0)
            {
                tail.Space = SpaceProcessingModeValues.Preserve;
            }

            run.InsertAfter(tail, text);
            SplitRunBefore(run, tail);
            return TopLevel(host, run);
        }

        // Offset is at (or past) the end of the paragraph's text.
        return lastTopLevelBefore ?? host.ChildElements.LastOrDefault(_ => _ is not ParagraphProperties);
    }

    static bool HasContentBefore(Run run, Text text)
    {
        foreach (var child in run.ChildElements)
        {
            if (child == text)
            {
                return false;
            }

            if (child is not RunProperties)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Moves <paramref name="from"/> and every following sibling out of <paramref name="run"/>
    /// into a fresh run (cloning run properties) inserted immediately after, so a paragraph-level
    /// boundary exists just before <paramref name="from"/>. When the run is nested (inside a
    /// hyperlink etc.) the split happens within that container — the boundary is then not at
    /// paragraph level, which is fine: token scanning never produces tokens straddling such
    /// containers and the host paragraphs Parchment mutates keep runs as direct children.
    /// </summary>
    static void SplitRunBefore(Run run, OpenXmlElement from)
    {
        var second = new Run();
        if (run.RunProperties != null)
        {
            second.AppendChild((RunProperties)run.RunProperties.CloneNode(true));
        }

        var moving = new List<OpenXmlElement>();
        var found = false;
        foreach (var child in run.ChildElements)
        {
            if (child == from)
            {
                found = true;
            }

            if (found)
            {
                moving.Add(child);
            }
        }

        foreach (var element in moving)
        {
            element.Remove();
            second.AppendChild(element);
        }

        run.Parent!.InsertAfter(second, run);
    }

    static void RemoveCoveredTexts(Paragraph host, int from, int to)
    {
        var consumed = 0;
        var toRemove = new List<Text>();
        foreach (var text in host.Descendants<Text>().ToList())
        {
            var start = consumed;
            var end = consumed + text.Text.Length;
            consumed = end;

            // Boundary splits already ran, so covered texts are whole texts.
            if (start >= from && end <= to && text.Text.Length > 0)
            {
                toRemove.Add(text);
            }
            else if (start >= to)
            {
                break;
            }
        }

        var owners = new HashSet<Run>();
        foreach (var text in toRemove)
        {
            if (text.Parent is Run run)
            {
                owners.Add(run);
            }

            text.Remove();
        }

        foreach (var run in owners)
        {
            if (!run.Descendants<Text>().Any())
            {
                run.Remove();
            }
        }
    }

    static OpenXmlElement TopLevel(Paragraph host, OpenXmlElement element)
    {
        var current = element;
        while (current.Parent != null &&
               current.Parent != host)
        {
            current = current.Parent;
        }

        return current;
    }
}
