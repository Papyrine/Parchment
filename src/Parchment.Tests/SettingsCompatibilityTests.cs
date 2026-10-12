using DocumentFormat.OpenXml.Validation;

/// <summary>
/// <c>w:settings</c> is a strict sequence with <c>w:compat</c> in the middle of it, so the
/// compatibility block a template lacks has to be created between whatever the settings already
/// hold, not at the end.
/// </summary>
// ReSharper disable PartialTypeWithSinglePart
public partial class SettingsCompatibilityTests
{
    const string namespaces =
        """
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        """;

    // CT_Settings members ahead of w:compat, in schema order (ECMA-376 §17.15.1.78).
    static readonly string[] preceding =
    [
        "w:writeProtection",
        "w:view",
        "w:zoom",
        "w:removePersonalInformation",
        "w:removeDateAndTime",
        "w:doNotDisplayPageBoundaries",
        "w:displayBackgroundShape",
        "w:printPostScriptOverText",
        "w:printFractionalCharacterWidth",
        "w:printFormsData",
        "w:embedTrueTypeFonts",
        "w:embedSystemFonts",
        "w:saveSubsetFonts",
        "w:saveFormsData",
        "w:mirrorMargins",
        "w:alignBordersAndEdges",
        "w:bordersDoNotSurroundHeader",
        "w:bordersDoNotSurroundFooter",
        "w:gutterAtTop",
        "w:hideSpellingErrors",
        "w:hideGrammaticalErrors",
        "w:activeWritingStyle",
        "w:proofState",
        "w:formsDesign",
        "w:attachedTemplate",
        "w:linkStyles",
        "w:stylePaneFormatFilter",
        "w:stylePaneSortMethod",
        "w:documentType",
        "w:mailMerge",
        "w:revisionView",
        "w:trackRevisions",
        "w:doNotTrackMoves",
        "w:doNotTrackFormatting",
        "w:documentProtection",
        "w:autoFormatOverride",
        "w:styleLockTheme",
        "w:styleLockQFSet",
        "w:defaultTabStop",
        "w:autoHyphenation",
        "w:consecutiveHyphenLimit",
        "w:hyphenationZone",
        "w:doNotHyphenateCaps",
        "w:showEnvelope",
        "w:summaryLength",
        "w:clickAndTypeStyle",
        "w:defaultTableStyle",
        "w:evenAndOddHeaders",
        "w:bookFoldRevPrinting",
        "w:bookFoldPrinting",
        "w:bookFoldPrintingSheets",
        "w:drawingGridHorizontalSpacing",
        "w:drawingGridVerticalSpacing",
        "w:displayHorizontalDrawingGridEvery",
        "w:displayVerticalDrawingGridEvery",
        "w:doNotUseMarginsForDrawingGridOrigin",
        "w:drawingGridHorizontalOrigin",
        "w:drawingGridVerticalOrigin",
        "w:doNotShadeFormData",
        "w:noPunctuationKerning",
        "w:characterSpacingControl",
        "w:printTwoOnOne",
        "w:strictFirstAndLastChars",
        "w:noLineBreaksAfter",
        "w:noLineBreaksBefore",
        "w:savePreviewPicture",
        "w:doNotValidateAgainstSchema",
        "w:saveInvalidXml",
        "w:ignoreMixedContent",
        "w:alwaysShowPlaceholderText",
        "w:doNotDemarcateInvalidXml",
        "w:saveXmlDataOnly",
        "w:useXSLTWhenSaving",
        "w:saveThroughXslt",
        "w:showXMLTags",
        "w:alwaysMergeEmptyNamespace",
        "w:updateFields",
        "w:hdrShapeDefaults",
        "w:footnotePr",
        "w:endnotePr"
    ];

    // CT_Settings members behind w:compat, in schema order, then the Word 2010 / 2013 extensions.
    static readonly string[] following =
    [
        "w:docVars",
        "w:rsids",
        "m:mathPr",
        "w:attachedSchema",
        "w:themeFontLang",
        "w:clrSchemeMapping",
        "w:doNotIncludeSubdocsInStats",
        "w:doNotAutoCompressPictures",
        "w:forceUpgrade",
        "w:captions",
        "w:readModeInkLockDown",
        "w:smartTagType",
        "sl:schemaLibrary",
        "w:shapeDefaults",
        "w:doNotEmbedSmartTags",
        "w:decimalSymbol",
        "w:listSeparator",
        "w14:docId",
        "w14:discardImageEditingData",
        "w14:defaultImageDpi",
        "w14:conflictMode",
        "w15:chartTrackingRefBased",
        "w15:docId"
    ];

    [Test]
    public async Task LandsAfterEachPrecedingMember()
    {
        foreach (var name in preceding)
        {
            await Assert.That(Apply(name)).IsEqualTo($"{name} w:compat");
        }
    }

    [Test]
    public async Task LandsBeforeEachFollowingMember()
    {
        foreach (var name in following)
        {
            await Assert.That(Apply(name)).IsEqualTo($"w:compat {name}");
        }
    }

    [Test]
    public async Task LandsBetweenEveryMemberAtOnce() =>
        await Assert.That(Apply([.. preceding, .. following]))
            .IsEqualTo(string.Join(' ', [.. preceding, "w:compat", .. following]));

    // w:smartTagType is in the CT_Settings sequence, behind w:compat, but the SDK has no class for
    // it: it loads as an unknown element, which the SDK's own ordered insert does not look at.
    [Test]
    public async Task LandsBeforeAMemberTheSdkDoesNotKnow()
    {
        await Assert.That(Apply("w:zoom", "w:smartTagType"))
            .IsEqualTo("w:zoom w:compat w:smartTagType");
        await Assert.That(Apply("w:zoom", "w:smartTagType", "w:smartTagType", "w:decimalSymbol"))
            .IsEqualTo("w:zoom w:compat w:smartTagType w:smartTagType w:decimalSymbol");
        await Assert.That(Apply("w:zoom", "w:docVars", "w:smartTagType"))
            .IsEqualTo("w:zoom w:compat w:docVars w:smartTagType");
    }

    [Test]
    public async Task LandsFirstInEmptySettings() =>
        await Assert.That(Apply()).IsEqualTo("w:compat");

    // The shape that was reported: document variables and no compatibility block.
    [Test]
    public async Task TemplateWithDocumentVariablesRendersValid()
    {
        using var template = new MemoryStream();
        using (var document = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new(new Body(new Paragraph(new Run(new Text("{{ Name }}")))));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new(
                new Zoom
                {
                    Percent = "100"
                },
                new DocumentVariables(
                    new DocumentVariable
                    {
                        Name = "a",
                        Val = "b"
                    }),
                new ShapeDefaults(),
                new DecimalSymbol
                {
                    Val = "."
                },
                new ListSeparator
                {
                    Val = ","
                });
        }

        template.Position = 0;
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        using var stream = new MemoryStream();
        await store.Render(
            new Model
            {
                Name = "x"
            },
            stream);

        stream.Position = 0;
        using var rendered = WordprocessingDocument.Open(stream, false);
        var settings = rendered.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        await Assert.That(string.Join(' ', settings.ChildElements.Select(_ => _.LocalName)))
            .IsEqualTo("zoom compat docVars shapeDefaults decimalSymbol listSeparator");
        await Assert.That(Errors(rendered)).IsEmpty();
    }

    [ParchmentBindable]
    public partial class Model
    {
        public required string Name;
    }

    // Runs SettingsCompatibility over a settings part holding the named members and returns the
    // resulting child order. The members are bare elements, so the validator's complaints about
    // their own missing attributes are left out: only the order is under test.
    static string Apply(params string[] names)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new(new Body());
        var children = string.Concat(names.Select(_ => $"<{_}/>"));
        mainPart.AddNewPart<DocumentSettingsPart>().Settings = new($"<w:settings {namespaces}>{children}</w:settings>");

        SettingsCompatibility.Apply(mainPart);

        var settings = mainPart.DocumentSettingsPart!.Settings!;
        var misplaced = Errors(document)
            .Where(_ => _.Contains("unexpected child element"))
            .ToList();
        // The validator rejects w:smartTagType itself wherever it sits, and then reports unrelated
        // members as unexpected, so with it present only the returned order is checked.
        if (misplaced.Count != 0 &&
            !names.Contains("w:smartTagType"))
        {
            throw new($"{string.Join(", ", names)}: {string.Join("; ", misplaced)}");
        }

        return string.Join(' ', settings.ChildElements.Select(_ => $"{_.Prefix}:{_.LocalName}"));
    }

    static List<string> Errors(WordprocessingDocument document) =>
        new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
}
