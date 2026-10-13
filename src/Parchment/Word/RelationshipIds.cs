using System.Runtime.CompilerServices;

/// <summary>
/// Picks the relationship id for every part and relationship Parchment adds. Adding one without an
/// id makes the Open XML SDK mint a random one (<c>R</c> + 16 hex digits), which lands in the
/// <c>.rels</c> file and in the XML referencing it, so two renders of the same input differ.
/// </summary>
/// <remarks>
/// An id only has to be unique within the part (or package) that owns the relationship, and the ids
/// a template arrives with are arbitrary, so each pick is checked against what that owner already
/// holds rather than assumed free. For every kind but a link that means reading the owner per call.
/// A document can hold thousands of links, and reading the part for each made them quadratic, so
/// the links of a part are indexed instead: target to id, and the highest counter taken. OpenXmlHtml,
/// Excelsior and caller delegates add relationships to the same parts without coming through here,
/// and an index does not see those, so <see cref="Forget"/> is called at each place a part is
/// handed to them and the next link brings the index up to date first. Every hyperlink
/// relationship Parchment adds itself has to be added here, and nothing else may add or remove a
/// relationship whose id starts with the link prefix without calling <see cref="Forget"/>, or
/// the index and the part disagree.
/// </remarks>
static class RelationshipIds
{
    const string linkPrefix = "rIdLink";

    // One entry per open package, so per render, holding an index per part that has been asked for
    // a link. A render owns its package and is the only reader and writer of its entry.
    static readonly ConditionalWeakTable<OpenXmlPackage, Dictionary<OpenXmlPart, LinkIndex>> linkIndexes = new();

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
    /// has none. A target linked several times shares one relationship. False, with nothing added,
    /// when <paramref name="url"/> is not something a relationship can point at.
    /// </summary>
    /// <remarks>
    /// Only an absolute url qualifies, the rule OpenXmlHtml applies to an <c>&lt;a href&gt;</c>
    /// (<c>LinkTarget.TryParseAbsolute</c>), so a markdown link and an html one agree. A url is as
    /// often a bound value as template text, and one that does not parse (<c>http://</c>, a space in
    /// the host) used to reach <c>new Uri</c> and throw <see cref="UriFormatException"/> out of the
    /// render. A relative one has nothing to be relative to in a generated document. A url that
    /// starts with a slash is tested by hand because <see cref="Uri.TryCreate(string, UriKind, out Uri)"/>
    /// accepts it as a file path on some platforms (<c>//host/page</c> on Windows, <c>/docs/page</c>
    /// elsewhere), which would make the output depend on the machine.
    /// </remarks>
    public static bool TryHyperlink(OpenXmlPart owner, string url, [NotNullWhen(true)] out string? id)
    {
        id = null;
        if (url.Length == 0 ||
            url[0] is '/' or '\\' ||
            !Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            return false;
        }

        id = ExternalHyperlink(owner, uri, url);
        return true;
    }

    /// <summary>
    /// The id, on <paramref name="owner"/>, of a hyperlink relationship another part holds: an
    /// external one shares the relationship the owner has to the same target, as in
    /// <see cref="TryHyperlink"/>, and any other is added under the next id.
    /// </summary>
    public static string Hyperlink(OpenXmlPart owner, HyperlinkRelationship link)
    {
        if (link.IsExternal)
        {
            // Already a relationship, so the target is one a relationship can hold: it is carried
            // over as it stands.
            return ExternalHyperlink(owner, link.Uri, link.Uri.OriginalString);
        }

        return Add(owner, LinkIndexFor(owner), link.Uri, false);
    }

    /// <summary>
    /// Has the hyperlink relationships of every part in the package of <paramref name="part"/> read
    /// again before the next link. Called wherever a part is handed to code that can add or remove
    /// relationships without going through this class.
    /// </summary>
    public static void Forget(OpenXmlPart part)
    {
        if (!linkIndexes.TryGetValue(part.OpenXmlPackage, out var indexes))
        {
            return;
        }

        foreach (var index in indexes.Values)
        {
            index.Current = false;
        }
    }

    static string ExternalHyperlink(OpenXmlPart owner, Uri uri, string target)
    {
        var index = LinkIndexFor(owner);
        if (index.External.TryGetValue(target, out var id))
        {
            return id;
        }

        return Add(owner, index, uri, true);
    }

    static string Add(OpenXmlPart owner, LinkIndex index, Uri uri, bool isExternal)
    {
        index.Highest++;
        var id = Format(linkPrefix, index.Highest);
        Record(index, owner.AddHyperlinkRelationship(uri, isExternal, id));
        return id;
    }

    static LinkIndex LinkIndexFor(OpenXmlPart owner)
    {
        var indexes = linkIndexes.GetValue(owner.OpenXmlPackage, _ => new(ReferenceEqualityComparer.Instance));
        if (!indexes.TryGetValue(owner, out var index))
        {
            index = new();
            indexes.Add(owner, index);
        }

        if (!index.Current)
        {
            Read(owner, index);
        }

        return index;
    }

    // Brings the index up to date with the part. What was added behind it is nearly always added
    // after everything it already knew, so the relationships it has seen are only checked to be the
    // same objects in the same order, which costs neither a string comparison nor an allocation,
    // and the ones after them are recorded. A relationship is immutable, so one that is still in
    // its place is still what the index holds for it. Anything else - one removed, or the order
    // changed - and the part is read from the start.
    static void Read(OpenXmlPart owner, LinkIndex index)
    {
        var seen = index.Seen;
        var known = seen.Count;
        var position = 0;
        var appended = true;
        foreach (var existing in owner.HyperlinkRelationships)
        {
            if (position >= known)
            {
                Record(index, existing);
            }
            else if (!ReferenceEquals(seen[position], existing))
            {
                appended = false;
                break;
            }

            position++;
        }

        if (!appended ||
            position < known)
        {
            seen.Clear();
            index.External.Clear();
            foreach (var existing in owner.HyperlinkRelationships)
            {
                Record(index, existing);
            }
        }

        index.Highest = Scan(owner, linkPrefix).highest;
        index.Current = true;
    }

    static void Record(LinkIndex index, HyperlinkRelationship relationship)
    {
        index.Seen.Add(relationship);
        if (relationship.IsExternal)
        {
            // The first relationship to a target is the one shared, as when the part was read in
            // order for every link.
            index.External.TryAdd(relationship.Uri.OriginalString, relationship.Id);
        }
    }

    sealed class LinkIndex
    {
        // The hyperlink relationships of the part, in the order it lists them.
        public List<HyperlinkRelationship> Seen { get; } = [];

        // External hyperlink target to relationship id.
        public Dictionary<string, string> External { get; } = new(StringComparer.Ordinal);

        // The highest counter any relationship of the part carries after the link prefix.
        public long Highest { get; set; }

        // Whether all of the above describe the part as it is now.
        public bool Current { get; set; }
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
