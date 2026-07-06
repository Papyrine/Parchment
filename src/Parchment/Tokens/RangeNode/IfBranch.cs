/// <summary>
/// One conditional branch of an <see cref="IfNode"/>. <paramref name="TagAnchorName"/> is the
/// anchor of the branch's opening tag paragraph (<c>{% if %}</c> for the first branch,
/// <c>{% elsif %}</c> for the rest) — the render-time boundary used to keep the chosen branch's
/// content positionally.
/// </summary>
sealed record IfBranch(
    Expression Condition,
    string TagAnchorName,
    IReadOnlyList<RangeNode> Body);
