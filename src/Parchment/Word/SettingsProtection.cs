/// <summary>
/// Writes <c>&lt;w:documentProtection w:edit="readOnly" w:enforcement="1"/&gt;</c> into the
/// template's settings part at registration time. Passwordless by design: protection is a
/// cooperative UI lock, not security — a password would add no real protection but would force
/// non-deterministic salt/hash generation into the output.
///
/// CT_Settings is a strict xsd:sequence, so the element cannot simply be appended — it must land
/// after every element that precedes <c>w:documentProtection</c> in the schema and before
/// everything else. Word rejects (or "repairs") files that get this wrong.
/// </summary>
static class SettingsProtection
{
    // Every CT_Settings member that precedes w:documentProtection, in schema order
    // (ECMA-376 §17.15.1.78). Insertion goes before the first child NOT in this set.
    static readonly HashSet<Type> preceding =
    [
        typeof(WriteProtection),
        typeof(View),
        typeof(Zoom),
        typeof(RemovePersonalInformation),
        typeof(RemoveDateAndTime),
        typeof(DoNotDisplayPageBoundaries),
        typeof(DisplayBackgroundShape),
        typeof(PrintPostScriptOverText),
        typeof(PrintFractionalCharacterWidth),
        typeof(PrintFormsData),
        typeof(EmbedTrueTypeFonts),
        typeof(EmbedSystemFonts),
        typeof(SaveSubsetFonts),
        typeof(SaveFormsData),
        typeof(MirrorMargins),
        typeof(AlignBorderAndEdges),
        typeof(BordersDoNotSurroundHeader),
        typeof(BordersDoNotSurroundFooter),
        typeof(GutterAtTop),
        typeof(HideSpellingErrors),
        typeof(HideGrammaticalErrors),
        typeof(ActiveWritingStyle),
        typeof(ProofState),
        typeof(FormsDesign),
        typeof(AttachedTemplate),
        typeof(LinkStyles),
        typeof(StylePaneFormatFilter),
        typeof(StylePaneSortMethods),
        typeof(DocumentType),
        typeof(MailMerge),
        typeof(RevisionView),
        typeof(TrackRevisions),
        typeof(DoNotTrackMoves),
        typeof(DoNotTrackFormatting)
    ];

    public static void Apply(MainDocumentPart mainPart)
    {
        var part = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        var settings = part.Settings ??= new();

        settings.RemoveAllChildren<DocumentProtection>();

        var protection = new DocumentProtection
        {
            Edit = DocumentProtectionValues.ReadOnly,
            Enforcement = true
        };

        var before = settings.ChildElements.FirstOrDefault(_ => !preceding.Contains(_.GetType()));
        if (before == null)
        {
            settings.AppendChild(protection);
        }
        else
        {
            settings.InsertBefore(protection, before);
        }
    }
}
