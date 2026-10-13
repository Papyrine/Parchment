/// <summary>
/// Guarantees the output declares <c>&lt;w:compatSetting w:name="compatibilityMode" w:val="15"/&gt;</c>
/// (Word 2013+). A docx built from scratch by the Open XML SDK — as Parchment does for code-authored
/// templates and the blank markdown host — omits the <c>w:compat</c> block entirely. Word treats a
/// MISSING compatibilityMode as mode 12 (Word 2007) and opens the file in "Compatibility Mode",
/// applying 2007-era layout rules (line-spacing tolerances, table cell spacing, list rendering).
/// Stamping mode 15 makes Word open the document normally, matching what Word itself writes on save.
///
/// Idempotent: does nothing when a compatibilityMode setting is already present (e.g. a Word-authored
/// template), so a mode the author deliberately chose is never overwritten.
///
/// CT_Settings is a strict xsd:sequence and <c>w:compat</c> sits in the middle of it: some twenty
/// members follow it (<c>w:docVars</c>, <c>w:rsids</c>, <c>m:mathPr</c>, <c>w:shapeDefaults</c>, the
/// <c>w14:</c> / <c>w15:</c> extension elements, …). A freshly created element is therefore placed
/// using the SDK's own knowledge of the sequence rather than a hand-kept list of the members that
/// follow.
/// </summary>
static class SettingsCompatibility
{
    const string wordUri = "http://schemas.microsoft.com/office/word";

    // Shared with SettingsProtection, which creates the same part at registration.
    internal const string RelationshipId = "rIdSettings";

    public static void Apply(MainDocumentPart mainPart)
    {
        // Pin the relationship id. AddNewPart<T>() with no id mints a random GUID-based rId, which
        // would make the rendered package non-deterministic (two renders of the same template would
        // differ, and a binary docx snapshot could never stabilise). A fixed id — matching the
        // rIdHeader / rIdFooter convention used elsewhere — keeps every render byte-identical.
        // RelationshipIds steps past it when the template already uses that id for something else.
        // This branch only runs for templates with no settings part of their own (code-authored docx
        // and the blank markdown host); a template that already has one keeps its existing part and id.
        var part = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>(RelationshipIds.Named(mainPart, RelationshipId));
        var settings = part.Settings ??= new();

        var compat = settings.GetFirstChild<Compatibility>();
        if (compat == null)
        {
            compat = new();
            Insert(settings, compat);
        }
        else if (HasCompatibilityMode(compat))
        {
            return;
        }

        // compatSetting is the trailing element inside CT_Compat (after the legacy boolean flags),
        // so appending is schema-safe regardless of what the block already contains.
        compat.AppendChild(
            new CompatibilitySetting
            {
                Name = CompatSettingNameValues.CompatibilityMode,
                Uri = wordUri,
                Val = "15"
            });
    }

    // The SDK knows the CT_Settings sequence, and AddChild inserts a child at its place in it, but
    // only reliably among children it has a class for: with an unknown element present it can put
    // the new child first or last. One member has no class (w:smartTagType), nor do extension
    // elements newer than the SDK. So the place is found in a scratch element holding empty copies
    // of the known children alone, and w:compat then goes straight after the last known child ahead
    // of it. That leaves every unknown element behind it, which is where the sequence has them:
    // w:smartTagType follows w:compat, and extension elements trail the whole sequence.
    static void Insert(Settings settings, Compatibility compat)
    {
        var known = settings.ChildElements
            .Where(_ => _ is not (OpenXmlUnknownElement or OpenXmlMiscNode or AlternateContent))
            .ToList();
        var scratch = new Settings();
        foreach (var child in known)
        {
            scratch.AppendChild(child.CloneNode(false));
        }

        var marker = new Compatibility();
        scratch.AddChild(marker);
        var ahead = scratch.ChildElements.TakeWhile(_ => _ != marker).Count();
        if (ahead == 0)
        {
            settings.PrependChild(compat);
            return;
        }

        settings.InsertAfter(compat, known[ahead - 1]);
    }

    // Read the name via InnerText: the SDK doesn't surface a strongly-typed enum member for every
    // compatSetting name, and InnerText is the raw attribute value regardless (same idiom Morph uses
    // to read the setting back).
    static bool HasCompatibilityMode(Compatibility compat) =>
        compat
            .Elements<CompatibilitySetting>()
            .Any(_ => string.Equals(_.Name?.InnerText, "compatibilityMode", StringComparison.OrdinalIgnoreCase));
}
