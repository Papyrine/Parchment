/// <summary>
/// Picks the relationship id for every part and relationship Parchment adds. Adding one without an
/// id makes the Open XML SDK mint a random one (<c>R</c> + 16 hex digits), which lands in the
/// <c>.rels</c> file and in the XML referencing it, so two renders of the same input differ.
/// </summary>
/// <remarks>
/// An id only has to be unique within the part (or package) that owns the relationship, and the ids
/// a template arrives with are arbitrary, so each pick is checked against what that owner already
/// holds rather than assumed free. The scan is repeated per call instead of cached: OpenXmlHtml and
/// Excelsior add relationships to the same parts, and a cache would not see those.
/// </remarks>
static class RelationshipIds
{
    /// <summary>
    /// The id for a part of which the owner holds at most one (settings, numbering, a properties
    /// part): <paramref name="name"/> itself, or <c>name2</c>, <c>name3</c>… when a template already
    /// uses it for something else.
    /// </summary>
    public static string Named(OpenXmlPartContainer owner, string name)
    {
        var (taken, highest) = Scan(owner, name);
        if (!taken)
        {
            return name;
        }

        return Format(name, Math.Max(highest, 1) + 1);
    }

    /// <summary>
    /// The id for one of many (links, images): <paramref name="prefix"/> followed by the counter
    /// after the highest the owner already holds, starting at 1.
    /// </summary>
    public static string Next(OpenXmlPartContainer owner, string prefix)
    {
        var (_, highest) = Scan(owner, prefix);
        return Format(prefix, highest + 1);
    }

    /// <summary>
    /// The id of an external hyperlink relationship to <paramref name="url"/>, added when the owner
    /// has none. A target linked several times shares one relationship.
    /// </summary>
    public static string Hyperlink(OpenXmlPartContainer owner, string url)
    {
        foreach (var existing in owner.HyperlinkRelationships)
        {
            if (existing.IsExternal &&
                string.Equals(existing.Uri.OriginalString, url, StringComparison.Ordinal))
            {
                return existing.Id;
            }
        }

        var id = Next(owner, "rIdLink");
        owner.AddHyperlinkRelationship(new(url, UriKind.RelativeOrAbsolute), true, id);
        return id;
    }

    static string Format(string prefix, long counter) =>
        prefix + counter.ToString(CultureInfo.InvariantCulture);

    static (bool taken, long highest) Scan(OpenXmlPartContainer owner, string prefix)
    {
        var taken = false;
        long highest = 0;

        void Visit(string id)
        {
            if (!id.StartsWith(prefix, StringComparison.Ordinal))
            {
                return;
            }

            if (id.Length == prefix.Length)
            {
                taken = true;
                return;
            }

            if (long.TryParse(id.AsSpan(prefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out var counter) &&
                counter > highest)
            {
                highest = counter;
            }
        }

        foreach (var pair in owner.Parts)
        {
            Visit(pair.RelationshipId);
        }

        foreach (var relationship in owner.ExternalRelationships)
        {
            Visit(relationship.Id);
        }

        foreach (var relationship in owner.HyperlinkRelationships)
        {
            Visit(relationship.Id);
        }

        foreach (var relationship in owner.DataPartReferenceRelationships)
        {
            Visit(relationship.Id);
        }

        return (taken, highest);
    }
}
