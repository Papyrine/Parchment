/// <summary>
/// Tells a token that reads a loop item from one that reads the root model.
/// </summary>
/// <remarks>
/// The attribute maps (<see cref="ExcelsiorTableMap"/>, <see cref="FormatMap"/>,
/// <see cref="StringListMap"/>, <see cref="EditableMap"/>) are keyed on dotted paths from the root
/// model, and a token is matched to them by its own dotted path. Inside
/// <c>{% for customer in Customers %}</c> the token <c>{{ customer.Notes }}</c> spells the same path
/// as a root <c>Customer.Notes</c>, but liquid resolves its first identifier to the loop variable.
/// Every consumer of those maps asks here first, so the item's member is never taken for the
/// root's.
/// </remarks>
static class LoopScope
{
    /// <summary>
    /// Whether <paramref name="site"/> starts with one of <paramref name="loopVariables"/>.
    /// Compared by exact spelling, which is how Fluid finds a loop variable. Model members are
    /// the ones matched without regard to case, so <c>{{ Customer.Notes }}</c> inside
    /// <c>{% for customer in Customers %}</c> is still the root's member.
    /// </summary>
    public static bool Shadows(IReadOnlyList<string> loopVariables, DocxTokenSite site)
    {
        if (loopVariables.Count == 0 ||
            site.References.Count == 0)
        {
            return false;
        }

        var root = site.References[0].Root;
        foreach (var loopVariable in loopVariables)
        {
            if (string.Equals(loopVariable, root, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Keeps <paramref name="loopVariables"/> current while paragraphs are walked in document
    /// order: a <c>{% for %}</c> opens its variable and the matching <c>{% endfor %}</c> closes it.
    /// Callers run after <see cref="ScopeTreeBuilder"/>, so the tags are known to balance.
    /// </summary>
    public static void Track(List<string> loopVariables, ParagraphClassification classification)
    {
        switch (classification.Block?.Kind)
        {
            case BlockTagKind.For:
                loopVariables.Add(classification.Block.LoopVariable!);
                break;
            case BlockTagKind.EndFor when loopVariables.Count > 0:
                loopVariables.RemoveAt(loopVariables.Count - 1);
                break;
        }
    }
}
