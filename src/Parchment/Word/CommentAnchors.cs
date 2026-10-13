using StringValue = DocumentFormat.OpenXml.StringValue;

/// <summary>
/// Keeps each comment on the first place a render anchors it.
/// </summary>
/// <remarks>
/// <para>
/// A comment is one note. Its text sits once in the comments part, and the document points at it
/// with a <c>w:commentRangeStart</c>, a <c>w:commentRangeEnd</c> and a <c>w:commentReference</c>
/// that carry its id. A <c>{% for %}</c> copies all three with its body, which leaves one comment
/// claiming a range per iteration: the ids repeat, and Word has one range to show.
/// </para>
/// <para>
/// The id cannot be changed on the copy the way <see cref="CloneIds"/> changes the others, because
/// here it is the reference. And the copy to keep is not known while the loop runs: the first
/// iteration can be the one whose <c>{% if %}</c> leaves the commented paragraph out. So this is
/// one pass over the part once everything in it has run, keeping the first of each anchor in
/// document order, and it runs only for a template that anchors a comment at all
/// (<see cref="TemplateIds.HoldsComments"/>). <see cref="ParagraphSplicer"/> needs no help from
/// it: a split puts each end of a range in the half it belongs to.
/// </para>
/// </remarks>
static class CommentAnchors
{
    public static void KeepFirst(OpenXmlCompositeElement root)
    {
        HashSet<string>? starts = null;
        HashSet<string>? ends = null;
        HashSet<string>? references = null;

        // Collected before removing: mutating the tree while enumerating Descendants is not safe.
        List<OpenXmlElement>? repeats = null;
        foreach (var element in root.Descendants())
        {
            var repeat = element switch
            {
                CommentRangeStart start => Repeats(ref starts, start.Id),
                CommentRangeEnd end => Repeats(ref ends, end.Id),
                CommentReference reference => Repeats(ref references, reference.Id),
                _ => false
            };
            if (repeat)
            {
                repeats ??= [];
                repeats.Add(element);
            }
        }

        if (repeats == null)
        {
            return;
        }

        foreach (var element in repeats)
        {
            var parent = element.Parent;
            element.Remove();

            // A reference sits in a run of its own, which goes with it.
            if (parent is Run run &&
                run.ChildElements.All(_ => _ is RunProperties))
            {
                run.Remove();
            }
        }
    }

    static bool Repeats(ref HashSet<string>? seen, StringValue? id)
    {
        if (id?.Value is not { } value)
        {
            return false;
        }

        seen ??= new(StringComparer.Ordinal);
        return !seen.Add(value);
    }
}
