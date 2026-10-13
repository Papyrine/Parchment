using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

/// <summary>
/// What a template already holds of the ids a copy of its content cannot share, read at
/// registration: where a render's <see cref="CloneIds"/> starts numbering from, and whether there
/// is anything for it or for <see cref="CommentAnchors"/> to do at all.
/// </summary>
/// <remarks>
/// Every part is read, with or without tokens: the ids are unique across the document, and a
/// render never opens a part that has no token. Written during registration only and read after
/// it, so the renders of a template share one.
/// </remarks>
sealed class TemplateIds
{
    /// <summary>
    /// The highest <c>wp:docPr</c> id.
    /// </summary>
    public uint MaxDrawing { get; private set; }

    /// <summary>
    /// The lowest content control id, or zero when none is below it.
    /// </summary>
    public int MinControl { get; private set; }

    /// <summary>
    /// The highest tracked-change id.
    /// </summary>
    public int MaxRevision { get; private set; }

    /// <summary>
    /// True when some part holds a drawing, a content control or a tracked change.
    /// </summary>
    public bool HoldsIds { get; private set; }

    /// <summary>
    /// True when some part anchors a comment.
    /// </summary>
    public bool HoldsComments { get; private set; }

    public void Read(OpenXmlCompositeElement root)
    {
        foreach (var element in root.Descendants())
        {
            if (element is DW.DocProperties drawing)
            {
                HoldsIds = true;
                if (uint.TryParse(drawing.Id?.InnerText, NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                {
                    MaxDrawing = Math.Max(MaxDrawing, id);
                }
            }
            else if (element is SdtId control)
            {
                HoldsIds = true;
                if (int.TryParse(control.Val?.InnerText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var id))
                {
                    MinControl = Math.Min(MinControl, id);
                }
            }
            else if (element is CommentRangeStart or CommentRangeEnd or CommentReference)
            {
                HoldsComments = true;
            }
            else if (CloneIds.RevisionId(element) is { } revision)
            {
                HoldsIds = true;
                if (int.TryParse(revision.Value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var id))
                {
                    MaxRevision = Math.Max(MaxRevision, id);
                }
            }
        }
    }
}
