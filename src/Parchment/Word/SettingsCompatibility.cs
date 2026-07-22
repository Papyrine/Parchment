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
/// CT_Settings is a strict xsd:sequence and <c>w:compat</c> sits near its end, so a freshly created
/// element is inserted before the first child that must follow it (<c>w:rsids</c>, <c>m:mathPr</c>, …)
/// and otherwise appended. In practice Parchment only creates the element when the settings part holds
/// at most <c>w:documentProtection</c> (which precedes <c>w:compat</c>), so the append path is the
/// common case; the following-set is a guard for richer settings parts.
/// </summary>
static class SettingsCompatibility
{
    const string wordUri = "http://schemas.microsoft.com/office/word";

    // CT_Settings members that FOLLOW w:compat, in schema order (ECMA-376 §17.15.1.78). A newly
    // created w:compat must land before the first of these. rsids and mathPr are the elements Word
    // almost always emits immediately after compat, so covering them handles real Word-authored
    // templates; the append fallback is correct whenever no post-compat sibling is present (the only
    // shape Parchment itself produces).
    static readonly HashSet<Type> following =
    [
        typeof(Rsids),
        typeof(DocumentFormat.OpenXml.Math.MathProperties),
        typeof(ThemeFontLanguages),
        typeof(ColorSchemeMapping),
        typeof(DecimalSymbol),
        typeof(ListSeparator)
    ];

    public static void Apply(MainDocumentPart mainPart)
    {
        // Pin the relationship id. AddNewPart<T>() with no id mints a random GUID-based rId, which
        // would make the rendered package non-deterministic (two renders of the same template would
        // differ, and a binary docx snapshot could never stabilise). A fixed id — matching the
        // rIdHeader / rIdFooter convention used elsewhere — keeps every render byte-identical.
        // This branch only runs for templates with no settings part of their own (code-authored docx
        // and the blank markdown host); a template that already has one keeps its existing part and id.
        var part = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>("rIdSettings");
        var settings = part.Settings ??= new();

        var compat = settings.GetFirstChild<Compatibility>();
        if (compat == null)
        {
            compat = new();
            var before = settings.ChildElements.FirstOrDefault(_ => following.Contains(_.GetType()));
            if (before == null)
            {
                settings.AppendChild(compat);
            }
            else
            {
                settings.InsertBefore(compat, before);
            }
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

    // Read the name via InnerText: the SDK doesn't surface a strongly-typed enum member for every
    // compatSetting name, and InnerText is the raw attribute value regardless (same idiom Morph uses
    // to read the setting back).
    static bool HasCompatibilityMode(Compatibility compat) =>
        compat
            .Elements<CompatibilitySetting>()
            .Any(_ => string.Equals(_.Name?.InnerText, "compatibilityMode", StringComparison.OrdinalIgnoreCase));
}
