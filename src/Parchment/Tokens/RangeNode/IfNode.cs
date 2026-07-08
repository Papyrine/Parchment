/// <summary>
/// <paramref name="ElseAnchorName"/> is the anchor of the <c>{% else %}</c> tag paragraph, or
/// null when the conditional has no else branch. Distinct from <c>ElseBody.Count == 0</c>: an
/// else branch whose content is purely static paragraphs has an empty body (static paragraphs
/// carry no anchors and no tree nodes) but must still render its physical range.
/// </summary>
sealed record IfNode(
    string OpenAnchorName,
    string CloseAnchorName,
    IReadOnlyList<IfBranch> Branches,
    string? ElseAnchorName,
    IReadOnlyList<RangeNode> ElseBody) :
    RangeNode;
