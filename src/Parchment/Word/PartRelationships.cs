/// <summary>
/// Gives the part a token sits in the relationships its produced content refers to. One per part
/// per render, shared by the runners of that part's loops and branches.
/// </summary>
/// <remarks>
/// <para>
/// A relationship id resolves against the part holding the element that carries it, so a link or
/// an image produced into a header has to hang off the header part. What Parchment adds itself
/// goes straight to <see cref="Owner"/>. OpenXmlHtml and Excelsior cannot do that: both take a
/// <see cref="MainDocumentPart"/> and add to it, and so does an <see cref="OpenXmlToken"/> delegate
/// handing them <see cref="IOpenXmlContext.MainPart"/>. Left there, the header keeps an id that
/// resolves nowhere and Word reports the file as corrupt. So outside the body, a relationship the
/// main part gained while a token was producing, and that the produced content refers to, is moved
/// across and the references pointed at its new id.
/// </para>
/// <para>
/// Only what the main part gained: an id it already held is one the content cannot be told to mean,
/// since <see cref="IOpenXmlContext.AddImagePart"/> hands out ids of the owner and the two parts
/// number independently. In the body the owner is the main part and nothing here does any work.
/// </para>
/// </remarks>
sealed class PartRelationships(MainDocumentPart mainPart, OpenXmlPart owner)
{
    const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // What has been moved, as main part id to owner id. The main part keeps its copy until
    // Complete: OpenXmlHtml numbers an image, and the drawing showing it, one past the highest
    // rImage{n} the main part holds, so dropping a header's first image at once would hand its
    // number, and with it its wp:docPr id, to the second.
    Dictionary<string, string>? moved;

    /// <summary>
    /// The part the content lands in, and so the one a relationship it refers to is added to.
    /// </summary>
    public OpenXmlPart Owner => owner;

    /// <summary>
    /// The ids the main part holds before a token produces, for a claim to compare against. Null
    /// in the body.
    /// </summary>
    public HashSet<string>? Snapshot()
    {
        if (ReferenceEquals(owner, mainPart))
        {
            return null;
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in mainPart.Parts)
        {
            ids.Add(pair.RelationshipId);
        }

        foreach (var relationship in mainPart.ExternalRelationships)
        {
            ids.Add(relationship.Id);
        }

        foreach (var relationship in mainPart.HyperlinkRelationships)
        {
            ids.Add(relationship.Id);
        }

        return ids;
    }

    /// <summary>
    /// Moves to the owner each relationship <paramref name="produced"/> refers to that the main
    /// part did not hold at <paramref name="before"/>, and points the references at its new id.
    /// </summary>
    public void Claim(IEnumerable<OpenXmlElement> produced, HashSet<string>? before)
    {
        if (before == null)
        {
            return;
        }

        foreach (var root in produced)
        {
            Repoint(root, before);
            foreach (var element in root.Descendants())
            {
                Repoint(element, before);
            }
        }
    }

    /// <summary>
    /// The same for a delegate that edited <paramref name="host"/> in place. The elements it found
    /// there, <paramref name="existing"/>, are the template's and already refer to the owner, so
    /// only what it added is looked at.
    /// </summary>
    public void Claim(OpenXmlElement host, HashSet<OpenXmlElement>? existing, HashSet<string>? before)
    {
        if (before == null ||
            existing == null)
        {
            return;
        }

        foreach (var element in host.Descendants())
        {
            if (!existing.Contains(element))
            {
                Repoint(element, before);
            }
        }
    }

    /// <summary>
    /// Takes what was moved off the main part, once the whole part has been rendered.
    /// </summary>
    public void Complete()
    {
        if (moved == null)
        {
            return;
        }

        foreach (var id in moved.Keys)
        {
            if (mainPart.TryGetPartById(id, out _))
            {
                // The relationship only: the owner holds the part as well now, so it stays.
                mainPart.DeletePart(id);
                continue;
            }

            mainPart.DeleteReferenceRelationship(id);
        }

        moved = null;
    }

    void Repoint(OpenXmlElement element, HashSet<string> before)
    {
        if (!element.HasAttributes)
        {
            return;
        }

        // A copy, so setting an attribute does not disturb the walk.
        foreach (var attribute in element.GetAttributes())
        {
            if (attribute.NamespaceUri != relationshipNamespace ||
                attribute.Value is not { Length: > 0 } id)
            {
                continue;
            }

            if (!TryMove(id, before, out var ownerId))
            {
                continue;
            }

            element.SetAttribute(new(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, ownerId));
        }
    }

    bool TryMove(string id, HashSet<string> before, [NotNullWhen(true)] out string? ownerId)
    {
        // Still on the main part, so later content may name it again - a delegate that looks for
        // a relationship to its target before adding one will find this.
        if (moved != null &&
            moved.TryGetValue(id, out ownerId))
        {
            return true;
        }

        if (before.Contains(id) ||
            !TryAddToOwner(id, out ownerId))
        {
            ownerId = null;
            return false;
        }

        moved ??= new(StringComparer.Ordinal);
        moved[id] = ownerId;
        return true;
    }

    bool TryAddToOwner(string id, [NotNullWhen(true)] out string? ownerId)
    {
        if (mainPart.TryGetPartById(id, out var part))
        {
            ownerId = RelationshipIds.Next(owner, Prefix(part));
            owner.AddPart(part, ownerId);
            return true;
        }

        foreach (var link in mainPart.HyperlinkRelationships)
        {
            if (link.Id == id)
            {
                ownerId = AddHyperlink(link);
                return true;
            }
        }

        foreach (var external in mainPart.ExternalRelationships)
        {
            if (external.Id == id)
            {
                ownerId = RelationshipIds.Next(owner, "rIdExternal");
                owner.AddExternalRelationship(external.RelationshipType, external.Uri, ownerId);
                return true;
            }
        }

        ownerId = null;
        return false;
    }

    string AddHyperlink(HyperlinkRelationship link)
    {
        if (link.IsExternal)
        {
            // Already a relationship on the main part, so the target is one a relationship can
            // hold: it is carried over as it stands, and shared when the owner points there too.
            foreach (var existing in owner.HyperlinkRelationships)
            {
                if (existing.IsExternal &&
                    string.Equals(existing.Uri.OriginalString, link.Uri.OriginalString, StringComparison.Ordinal))
                {
                    return existing.Id;
                }
            }

            var externalId = RelationshipIds.Next(owner, "rIdLink");
            owner.AddHyperlinkRelationship(link.Uri, true, externalId);
            return externalId;
        }

        var id = RelationshipIds.Next(owner, "rIdLink");
        owner.AddHyperlinkRelationship(link.Uri, false, id);
        return id;
    }

    static string Prefix(OpenXmlPart part)
    {
        if (part is ImagePart)
        {
            return "rIdImage";
        }

        return "rIdPart";
    }
}
