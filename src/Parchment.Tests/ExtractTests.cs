using System.Globalization;
using DocumentFormat.OpenXml.Validation;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
// ReSharper disable PartialTypeWithSinglePart

public partial class ExtractTests
{
    // Deliberately NOT [ParchmentBindable]: the generator never sees it, so its editable map is
    // empty for the wrong reason — the guard has to say "not generated", not "no fields".
    public class NeverGeneratedModel
    {
        [EditableField]
        public string PurchaseOrder { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class NoFieldsModel
    {
        public string Title { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class OffsetFormatsModel
    {
        [EditableField(DateFormat = "yyyy-MM-dd HH:mm")]
        public DateTimeOffset NoOffset { get; set; }

        [EditableField(DateFormat = "HH:mm")]
        public DateTimeOffset NoDate { get; set; }

        [EditableField(DateFormat = "dd/MM/yy")]
        public DateTimeOffset ShortYear { get; set; }
    }

    [Test]
    public async Task ExtractingANeverGeneratedModelSaysSo()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        var exception = await Assert.That(
                () => ParchmentExtractor.Extract<NeverGeneratedModel>(stream))
            .Throws<ParchmentExtractionException>();
        await Assert.That(exception!.Message).Contains("no pre-compiled Parchment accessors");
        await Assert.That(exception.Message).Contains("ParchmentBindable");
    }

    [Test]
    public async Task ExtractingAModelWithoutEditableFieldsSaysSo()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        var exception = await Assert.That(
                () => ParchmentExtractor.Extract<NoFieldsModel>(stream))
            .Throws<ParchmentExtractionException>();
        await Assert.That(exception!.Message).Contains("declares no [EditableField] members");
    }

    [Test]
    public async Task NoEditRoundTrip()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(Field(result, "PurchaseOrder").Value).IsEqualTo("PO-2026-17");
        await Assert.That((bool)Field(result, "Approved").Value!).IsTrue();
        await Assert.That(Field(result, "Delivery").Value).IsEqualTo(new Date(2026, 7, 6));
        await Assert.That(Field(result, "Status").Value).IsEqualTo(EditableFieldTests.QuoteStatus.Submitted);
        await Assert.That(Field(result, "Discount").Value).IsEqualTo(10m);
        await Assert.That(Field(result, "Notes").State).IsEqualTo(FieldState.Empty);
        await Assert.That(Field(result, "Instructions").Value).IsEqualTo("Line one\nLine two");

        // Apply onto a model with divergent editable values — every editable member converges
        // on the document's values, including null for the placeholder field.
        var target = EditableFieldTests.NewOrder();
        target.PurchaseOrder = "OTHER";
        target.Approved = false;
        target.Delivery = new(2000, 1, 1);
        target.Status = EditableFieldTests.QuoteStatus.Draft;
        target.Discount = 0m;
        target.Notes = "prefilled";
        target.Instructions = null;

        result.ApplyTo(target);

        await Assert.That(target.PurchaseOrder).IsEqualTo("PO-2026-17");
        await Assert.That(target.Approved).IsTrue();
        await Assert.That(target.Delivery).IsEqualTo(new(2026, 7, 6));
        await Assert.That(target.Status).IsEqualTo(EditableFieldTests.QuoteStatus.Submitted);
        await Assert.That(target.Discount).IsEqualTo(10m);
        await Assert.That(target.Notes).IsNull();
        await Assert.That(target.Instructions).IsEqualTo("Line one\nLine two");
    }

    [Test]
    public async Task EditedValuesRoundTrip()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body =>
        {
            EditSdtText(body, "PurchaseOrder", "PO-EDITED");
            SetCheckbox(body, "Approved", false);
            SetFullDate(body, "Delivery", new(2026, 8, 1));
            EditSdtText(body, "Status", "Accepted");
            EditSdtText(body, "Discount", "12.5");
            EditSdtText(body, "Notes", "call before delivery");
            EditSdtText(body, "Instructions", "single line");
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(result.AllExtracted).IsTrue();

        var target = EditableFieldTests.NewOrder();
        result.ApplyTo(target);

        await Assert.That(target.PurchaseOrder).IsEqualTo("PO-EDITED");
        await Assert.That(target.Approved).IsFalse();
        await Assert.That(target.Delivery).IsEqualTo(new(2026, 8, 1));
        await Assert.That(target.Status).IsEqualTo(EditableFieldTests.QuoteStatus.Accepted);
        await Assert.That(target.Discount).IsEqualTo(12.5m);
        await Assert.That(target.Notes).IsEqualTo("call before delivery");
        await Assert.That(target.Instructions).IsEqualTo("single line");
    }

    [Test]
    public async Task FormattingOnPlainFieldDoesNotLeakIntoValue()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        // A plain-text control sits inside the editable range, so Word lets a user bold/italicise it.
        // That formatting is cosmetic — extraction reads w:t only and must ignore w:rPr, so the value
        // comes back identical to unformatted text. Split across two runs to prove concatenation too.
        Edit(stream, body =>
        {
            var sdt = FindSdt(body, "PurchaseOrder");
            sdt.SdtProperties!.RemoveAllChildren<ShowingPlaceholder>();
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            content.RemoveAllChildren();
            content.AppendChild(new Run(new RunProperties(new Bold()), new Text("PO-")));
            content.AppendChild(new Run(new RunProperties(new Italic()), new Text("BOLD")));
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "PurchaseOrder").Value).IsEqualTo("PO-BOLD");
    }

    // The plain-text reader and the rich-text one share EditableContent, so the wrappers Word
    // writes around runs read the same way in both: tracked changes as if accepted, everything
    // else transparent, a twice-stored alternative once.
    [Test]
    [Arguments("""<w:r><w:t xml:space="preserve">keep </w:t></w:r><w:ins w:id="901" w:author="u"><w:r><w:t>typed</w:t></w:r></w:ins><w:del w:id="902" w:author="u"><w:r><w:delText>gone</w:delText></w:r></w:del>""", "keep typed")]
    [Arguments("""<w:moveFrom w:id="901" w:author="u"><w:r><w:t xml:space="preserve">moved </w:t></w:r></w:moveFrom><w:r><w:t xml:space="preserve">stay </w:t></w:r><w:moveTo w:id="902" w:author="u"><w:r><w:t>moved</w:t></w:r></w:moveTo>""", "stay moved")]
    [Arguments("""<w:smartTag w:uri="urn:x" w:element="place"><w:r><w:t>Paris</w:t></w:r></w:smartTag><w:fldSimple w:instr=" DATE "><w:r><w:t xml:space="preserve"> 2026</w:t></w:r></w:fldSimple>""", "Paris 2026")]
    [Arguments("""<w:r><w:t xml:space="preserve">a </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent><w:hyperlink w:anchor="bm"><w:r><w:t>b</w:t></w:r></w:hyperlink></w:sdtContent></w:sdt>""", "a b")]
    [Arguments("""<mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>new</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>old</w:t></w:r></mc:Fallback></mc:AlternateContent>""", "new")]
    [Arguments("""<w:r><w:t>x</w:t></w:r><m:oMath><m:r><m:t>=1</m:t></m:r></m:oMath>""", "x=1")]
    public async Task PlainFieldReadsThroughWhatWordWrapsRunsIn(string contentXml, string expected)
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body =>
        {
            var sdt = FindSdt(body, "PurchaseOrder");
            sdt.SdtProperties!.RemoveAllChildren<ShowingPlaceholder>();
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            sdt.ReplaceChild(
                new SdtContentRun(
                    $"""<w:sdtContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{contentXml}</w:sdtContent>"""),
                content);
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "PurchaseOrder").Value).IsEqualTo(expected);
    }

    [Test]
    public async Task TabSavedByWordAsAnElementIsRead()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        // The run shape Word leaves behind on saving a control whose value held a tab: the tab
        // becomes w:tab at the head of a run of its own, and the line break likewise.
        Edit(stream, body =>
        {
            SetSdtRuns(
                body,
                "PurchaseOrder",
                new Run(new Text("left")),
                new Run(new TabChar(), new Text("right")));
            SetSdtRuns(
                body,
                "Instructions",
                new Run(new Text("col1")),
                new Run(new TabChar(), new Text("col2")),
                new Run(new Break(), new Text("row2")),
                new Run(new TabChar(), new Text("cell")));
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "PurchaseOrder").Value).IsEqualTo("left\tright");
        await Assert.That(Field(result, "PurchaseOrder").RawText).IsEqualTo("left\tright");
        await Assert.That(Field(result, "Instructions").Value).IsEqualTo("col1\tcol2\nrow2\tcell");
    }

    [Test]
    public async Task RunContentWordWritesAsElementsIsRead()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body =>
        {
            SetSdtRuns(
                body,
                "PurchaseOrder",
                new Run(
                    new Text("co"),
                    new NoBreakHyphen(),
                    new Text("op"),
                    new SoftHyphen(),
                    new Text("s ")),
                new Run(
                    // A Unicode font: the code point is the character.
                    new SymbolChar
                    {
                        Font = "Arial",
                        Char = "2022"
                    },
                    // A symbol font: the private use code point Word stored, as stored.
                    new SymbolChar
                    {
                        Font = "Wingdings",
                        Char = "F0E0"
                    },
                    // Not a code point at all.
                    new SymbolChar
                    {
                        Font = "Arial",
                        Char = "zz"
                    }));
            SetSdtRuns(
                body,
                "Instructions",
                new Run(new Text("a"), new CarriageReturn(), new Text("b")));
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "PurchaseOrder").Value).IsEqualTo("co‑ops •");
        await Assert.That(Field(result, "Instructions").Value).IsEqualTo("a\nb");
    }

    [Test]
    public async Task TabInValueRendersAsATabElementAndRoundTrips()
    {
        var order = EditableFieldTests.NewOrder();
        order.PurchaseOrder = "left\tright";
        order.Instructions = "col1\tcol2\nrow2\tcell\t";
        using var stream = await RenderOrder(order);

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;

            // A tab inside w:t is whitespace to Word, which rewrites it as w:tab on save. Written
            // as the element up front, the control is the same before and after that save.
            await Assert.That(body.Descendants<Text>().Any(_ => _.Text.Contains('\t'))).IsFalse();
            await Assert.That(FindSdt(body, "PurchaseOrder").Descendants<TabChar>().Count()).IsEqualTo(1);
            await Assert.That(FindSdt(body, "Instructions").Descendants<TabChar>().Count()).IsEqualTo(3);

            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
                .Validate(doc)
                .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
                .ToList();
            await Assert.That(errors).IsEmpty();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(Field(result, "PurchaseOrder").Value).IsEqualTo("left\tright");
        await Assert.That(Field(result, "Instructions").Value).IsEqualTo("col1\tcol2\nrow2\tcell\t");
    }

    static void SetSdtRuns(Body body, string tag, params Run[] runs)
    {
        var sdt = FindSdt(body, tag);
        sdt.SdtProperties!.RemoveAllChildren<ShowingPlaceholder>();
        var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
        content.RemoveAllChildren();
        content.Append(runs);
    }

    [Test]
    public async Task DeletedControlReportsMissing()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => FindSdt(body, "PurchaseOrder").Remove());

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(result.AllExtracted).IsFalse();
        await Assert.That(Field(result, "PurchaseOrder").State).IsEqualTo(FieldState.Missing);

        var target = EditableFieldTests.NewOrder();
        target.PurchaseOrder = "KEEP";
        result.ApplyTo(target);
        await Assert.That(target.PurchaseOrder).IsEqualTo("KEEP");
    }

    [Test]
    public async Task UnparseableNumberReportsParseFailedWithRawText()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => EditSdtText(body, "Discount", "about twelve"));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        var discount = Field(result, "Discount");
        await Assert.That(result.AllExtracted).IsFalse();
        await Assert.That(discount.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(discount.RawText).IsEqualTo("about twelve");

        var target = EditableFieldTests.NewOrder();
        target.Discount = 99m;
        result.ApplyTo(target);
        await Assert.That(target.Discount).IsEqualTo(99m);
    }

    [Test]
    public async Task ClearedNumberIsEmptyAndSkippedForNonNullable()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => EditSdtText(body, "Discount", ""));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "Discount").State).IsEqualTo(FieldState.Empty);

        var target = EditableFieldTests.NewOrder();
        target.Discount = 99m;
        result.ApplyTo(target);
        await Assert.That(target.Discount).IsEqualTo(99m);
    }

    [Test]
    public async Task DuplicateTagIsReportedAndFirstOccurrenceWins()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body =>
        {
            var sdt = FindSdt(body, "PurchaseOrder");
            var clone = (SdtRun)sdt.CloneNode(true);
            var text = clone.Descendants<Text>().First();
            text.Text = "SECOND";
            sdt.Parent!.AppendChild(clone);
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        var occurrences = result.Fields.Where(_ => _.Path == "PurchaseOrder").ToList();
        await Assert.That(occurrences.Count).IsEqualTo(2);
        await Assert.That(occurrences[0].State).IsEqualTo(FieldState.Extracted);
        await Assert.That(occurrences[0].Value).IsEqualTo("PO-2026-17");
        await Assert.That(occurrences[1].State).IsEqualTo(FieldState.Duplicate);
        await Assert.That(occurrences[1].RawText).IsEqualTo("SECOND");
        await Assert.That(result.AllExtracted).IsFalse();
    }

    [Test]
    public async Task NumberParsingHonorsCulture()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => EditSdtText(body, "Discount", "12,5"));

        var german = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(
            stream,
            CultureInfo.GetCultureInfo("de-DE"));
        await Assert.That(Field(german, "Discount").Value).IsEqualTo(12.5m);
    }

    [Test]
    public async Task NestedPathAppliesAndNullIntermediateThrowsBeforeApplying()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Ref {{ Reference }}

            Email: {{ Customer.ContactEmail }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableFieldTests.EditableQuote>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new EditableFieldTests.EditableQuote
            {
                Reference = "Q-7",
                Customer = new()
                {
                    ContactEmail = "ada@example.com"
                }
            },
            stream);
        stream.Position = 0;

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableQuote>(stream);
        await Assert.That(result.AllExtracted).IsTrue();

        var good = new EditableFieldTests.EditableQuote
        {
            Reference = "OLD",
            Customer = new()
            {
                ContactEmail = "old@example.com"
            }
        };
        result.ApplyTo(good);
        await Assert.That(good.Reference).IsEqualTo("Q-7");
        await Assert.That(good.Customer.ContactEmail).IsEqualTo("ada@example.com");

        // Null intermediate: reachability is validated before anything is written, so the
        // reachable Reference field must remain untouched after the throw.
        var bad = new EditableFieldTests.EditableQuote
        {
            Reference = "KEEP",
            Customer = null!
        };
        var exception = await Assert.That(() => result.ApplyTo(bad))
            .Throws<ParchmentExtractionException>();
        await Assert.That(exception!.Message).Contains("Customer.ContactEmail");
        await Assert.That(bad.Reference).IsEqualTo("KEEP");
    }

    [Test]
    public async Task ModelWithoutEditableMembersIsRejected()
    {
        using var stream = new MemoryStream();
        var exception = await Assert.That(() => ParchmentExtractor.Extract<Invoice>(stream))
            .Throws<ParchmentExtractionException>();
        await Assert.That(exception!.Message).Contains("no [EditableField]");
    }

    [Test]
    public async Task NonDocxStreamIsRejected()
    {
        using var stream = new MemoryStream([1, 2, 3, 4]);
        await Assert.That(() => ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream))
            .Throws<ParchmentExtractionException>();
    }

    [Test]
    public async Task DocumentWithoutMatchingControlsReportsAllMissing()
    {
        using var foreign = DocxTemplateBuilder.Build("Unrelated document.");

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(foreign);

        await Assert.That(result.AllExtracted).IsFalse();
        await Assert.That(result.Fields.Count).IsEqualTo(10);
        foreach (var field in result.Fields)
        {
            await Assert.That(field.State).IsEqualTo(FieldState.Missing);
        }
    }

    [Test]
    public async Task TemporalTypesRoundTrip()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        await Assert.That(result.AllExtracted).IsTrue();

        // DateOnly — canonical w:fullDate at midnight.
        await Assert.That(Field(result, "Delivery").Value).IsEqualTo(new Date(2026, 7, 6));

        // DateTime — time-of-day preserved through w:fullDate; Kind normalizes to Unspecified.
        await Assert.That(Field(result, "DispatchedAt").Value).IsEqualTo(new DateTime(2026, 7, 6, 14, 30, 0));

        // DateTimeOffset — the +10:00 offset survives (parsed from the ISO run text, not a
        // zeroed fullDate). This is the case the old code corrupted to +00:00.
        var signed = (DateTimeOffset)Field(result, "SignedAt").Value!;
        await Assert.That(signed).IsEqualTo(new(2026, 7, 6, 9, 0, 0, TimeSpan.FromHours(10)));
        await Assert.That(signed.Offset).IsEqualTo(TimeSpan.FromHours(10));

        // TimeOnly — round-trips via plain text.
        await Assert.That(Field(result, "PickupTime").Value).IsEqualTo(new Time(16, 45, 0));

        var target = new EditableFieldTests.EditableOrder
        {
            Number = "x",
            Tags = [],
            IncludeNotes = false,
            PurchaseOrder = ""
        };
        result.ApplyTo(target);

        await Assert.That(target.DispatchedAt).IsEqualTo(new(2026, 7, 6, 14, 30, 0));
        await Assert.That(target.SignedAt).IsEqualTo(new(2026, 7, 6, 9, 0, 0, TimeSpan.FromHours(10)));
        await Assert.That(target.PickupTime).IsEqualTo(new(16, 45, 0));
    }

    [Test]
    public async Task EditedDateTimeOffsetPreservesNewOffset()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        // User types a different instant with a different offset into the plain-text control.
        Edit(stream, body => EditSdtText(body, "SignedAt", "2026-12-01T23:15:00-05:00"));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        var signed = (DateTimeOffset)Field(result, "SignedAt").Value!;

        await Assert.That(signed).IsEqualTo(new(2026, 12, 1, 23, 15, 0, TimeSpan.FromHours(-5)));
    }

    [Test]
    public async Task UnparseableTimeReportsParseFailed()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => EditSdtText(body, "PickupTime", "half four"));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        var pickup = Field(result, "PickupTime");

        await Assert.That(pickup.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(pickup.RawText).IsEqualTo("half four");
        await Assert.That(result.AllExtracted).IsFalse();
    }

    // w14:ST_OnOff spells each state two ways, and a producer other than Parchment may write either.
    [Test]
    [Arguments("1", true)]
    [Arguments("true", true)]
    [Arguments("0", false)]
    [Arguments("false", false)]
    public async Task CheckboxStateAcceptsEverySpelling(string val, bool expected)
    {
        var field = await ExtractApproved(!expected, _ => Checked(_).SetAttribute(CheckedVal(val)));

        await Assert.That(field.State).IsEqualTo(FieldState.Extracted);
        await Assert.That((bool) field.Value!).IsEqualTo(expected);
    }

    [Test]
    public async Task CheckboxStateWithoutValIsChecked()
    {
        var field = await ExtractApproved(false, _ => Checked(_).Val = null);

        await Assert.That(field.State).IsEqualTo(FieldState.Extracted);
        await Assert.That((bool) field.Value!).IsTrue();
    }

    [Test]
    public async Task CheckboxWithoutCheckedElementIsUnchecked()
    {
        var field = await ExtractApproved(true, _ => Checked(_).Remove());

        await Assert.That(field.State).IsEqualTo(FieldState.Extracted);
        await Assert.That((bool) field.Value!).IsFalse();
    }

    // "on" and "True" are legal for w:val but not for w14:val.
    [Test]
    [Arguments("on")]
    [Arguments("True")]
    [Arguments("")]
    public async Task CheckboxStateOutsideTheSchemaReportsParseFailed(string val)
    {
        var field = await ExtractApproved(true, _ => Checked(_).SetAttribute(CheckedVal(val)));

        await Assert.That(field.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(field.Value).IsNull();
        await Assert.That(field.RawText).IsEqualTo("☒");
    }

    // The glyph is display text a template can change, so it is never read as the state.
    [Test]
    public async Task ControlThatLostItsCheckboxReportsParseFailed()
    {
        var field = await ExtractApproved(true, _ => _.SdtProperties!.RemoveAllChildren<W14.SdtContentCheckBox>());

        await Assert.That(field.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(field.RawText).IsEqualTo("☒");
    }

    [Test]
    [Arguments("0")]
    [Arguments("false")]
    [Arguments("off")]
    public async Task PlaceholderFlagSwitchedOffIsNotEmpty(string val)
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body =>
        {
            EditSdtText(body, "PurchaseOrder", "PO-TYPED");
            FindSdt(body, "PurchaseOrder").SdtProperties!.AppendChild(
                new ShowingPlaceholder
                {
                    Val = new OnOffValue
                    {
                        InnerText = val
                    }
                });
        });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        var field = Field(result, "PurchaseOrder");

        await Assert.That(field.State).IsEqualTo(FieldState.Extracted);
        await Assert.That(field.Value).IsEqualTo("PO-TYPED");
    }

    static async Task<ExtractedField> ExtractApproved(bool approved, Action<SdtRun> edit)
    {
        var order = EditableFieldTests.NewOrder();
        order.Approved = approved;
        using var stream = await RenderOrder(order);

        Edit(stream, body => edit(FindSdt(body, "Approved")));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        return Field(result, "Approved");
    }

    static W14.Checked Checked(SdtRun sdt) =>
        sdt.SdtProperties!
            .GetFirstChild<W14.SdtContentCheckBox>()!
            .GetFirstChild<W14.Checked>()!;

    static OpenXmlAttribute CheckedVal(string val) =>
        new("w14", "val", "http://schemas.microsoft.com/office/word/2010/wordml", val);

    [Test]
    [Arguments("09:00")]
    [Arguments("09:00:00+10:00")]
    [Arguments("6 July 09:00")]
    [Arguments("07/06/26 09:00")]
    public async Task DateTimeOffsetTextWithoutAYearReportsParseFailed(string text)
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        // The framework parser would fill the gap with today's date (or the current year) and
        // the machine's offset, so the value would change with the day and machine of extraction.
        Edit(stream, body => EditSdtText(body, "SignedAt", text));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        var signed = Field(result, "SignedAt");

        await Assert.That(signed.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(signed.Value).IsNull();
        await Assert.That(signed.RawText).IsEqualTo(text);
    }

    [Test]
    [Arguments("2026-12-01T23:15:00")]
    [Arguments("2026-12-01 23:15")]
    [Arguments("1 December 2026 23:15")]
    public async Task DateTimeOffsetTextWithoutAnOffsetReadsAsZeroOffset(string text)
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => EditSdtText(body, "SignedAt", text));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        var signed = (DateTimeOffset)Field(result, "SignedAt").Value!;

        await Assert.That(signed.DateTime).IsEqualTo(new(2026, 12, 1, 23, 15, 0));
        await Assert.That(signed.Offset).IsEqualTo(TimeSpan.Zero);
    }

    [Test]
    public async Task DateTimeOffsetFormatsWithoutOffsetDateOrFullYear()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ NoOffset }}

            {{ NoDate }}

            {{ ShortYear }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<OffsetFormatsModel>(template);
        var instant = new DateTimeOffset(2026, 7, 6, 9, 0, 0, TimeSpan.FromHours(10));
        using var stream = new MemoryStream();
        await store.Render(
            new OffsetFormatsModel
            {
                NoOffset = instant,
                NoDate = instant,
                ShortYear = instant
            },
            stream);
        stream.Position = 0;

        var result = ParchmentExtractor.Extract<OffsetFormatsModel>(stream);

        // The format drops the offset, so it cannot come back: zero, never the machine's.
        var noOffset = (DateTimeOffset)Field(result, "NoOffset").Value!;
        await Assert.That(noOffset.DateTime).IsEqualTo(new(2026, 7, 6, 9, 0, 0));
        await Assert.That(noOffset.Offset).IsEqualTo(TimeSpan.Zero);

        // A format with no year has nothing to rebuild the date from.
        await Assert.That(Field(result, "NoDate").State).IsEqualTo(FieldState.ParseFailed);

        // A declared two digit year is read with the extraction culture's pivot.
        var shortYear = (DateTimeOffset)Field(result, "ShortYear").Value!;
        await Assert.That(shortYear.DateTime).IsEqualTo(new(2026, 7, 6));
        await Assert.That(shortYear.Offset).IsEqualTo(TimeSpan.Zero);
    }

    [Test]
    [Arguments("Delivery", "09:00")]
    [Arguments("Delivery", "6 July")]
    [Arguments("Delivery", "07/06/26")]
    [Arguments("DispatchedAt", "09:00")]
    [Arguments("DispatchedAt", "07-06 09:00")]
    public async Task DateTextWithoutAYearReportsParseFailed(string tag, string text)
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(
            stream,
            body =>
            {
                ClearFullDate(body, tag);
                EditSdtText(body, tag, text);
            });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);
        var field = Field(result, tag);

        await Assert.That(field.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(field.RawText).IsEqualTo(text);
    }

    [Test]
    public async Task DateTextIsReadAsTypedWhenFullDateIsAbsent()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(
            stream,
            body =>
            {
                ClearFullDate(body, "Delivery");
                ClearFullDate(body, "DispatchedAt");
                // A DateOnly given a date and time keeps the date.
                EditSdtText(body, "Delivery", "2027-01-31 16:45");
                // An offset in the text is not converted to the machine's zone.
                EditSdtText(body, "DispatchedAt", "2027-01-31T23:30:00-05:00");
            });

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "Delivery").Value).IsEqualTo(new Date(2027, 1, 31));
        var dispatched = (DateTime)Field(result, "DispatchedAt").Value!;
        await Assert.That(dispatched).IsEqualTo(new(2027, 1, 31, 23, 30, 0));
        await Assert.That(dispatched.Kind).IsEqualTo(DateTimeKind.Unspecified);
    }

    [Test]
    public async Task TimeGivenADateAndTimeReportsParseFailed()
    {
        using var stream = await RenderOrder(EditableFieldTests.NewOrder());

        Edit(stream, body => EditSdtText(body, "PickupTime", "2026-07-06 16:45:30"));

        var result = ParchmentExtractor.Extract<EditableFieldTests.EditableOrder>(stream);

        await Assert.That(Field(result, "PickupTime").State).IsEqualTo(FieldState.ParseFailed);
    }

    static async Task<MemoryStream> RenderOrder(EditableFieldTests.EditableOrder model)
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Order {{ Number }}

            {{ PurchaseOrder }}

            {{ Approved }}

            {{ Delivery }}

            {{ DispatchedAt }}

            {{ SignedAt }}

            {{ PickupTime }}

            {{ Status }}

            {{ Discount }}

            {{ Notes }}

            {{ Instructions }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableFieldTests.EditableOrder>(template);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static ExtractedField Field<TModel>(ExtractResult<TModel> result, string path) =>
        result.Fields.Single(_ => _.Path == path);

    /// <summary>
    /// Simulates a user editing the document: opens the stream writable, mutates, saves.
    /// </summary>
    static void Edit(MemoryStream stream, Action<Body> edit)
    {
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            edit(doc.MainDocumentPart!.Document!.Body!);
            doc.Save();
        }

        stream.Position = 0;
    }

    static SdtRun FindSdt(Body body, string tag)
    {
        var sdt = body.Descendants<SdtRun>()
            .FirstOrDefault(_ => _.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag);
        if (sdt == null)
        {
            throw new InvalidOperationException($"No sdt with tag '{tag}' found");
        }

        return sdt;
    }

    /// <summary>
    /// Replaces the control's content the way Word does when a user types: placeholder flag
    /// removed, content replaced with a plain run.
    /// </summary>
    static void EditSdtText(Body body, string tag, string text)
    {
        var sdt = FindSdt(body, tag);
        sdt.SdtProperties!.RemoveAllChildren<ShowingPlaceholder>();
        var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
        content.RemoveAllChildren();
        content.AppendChild(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
    }

    static void SetCheckbox(Body body, string tag, bool value)
    {
        var sdt = FindSdt(body, tag);
        var isChecked = sdt.SdtProperties!
            .GetFirstChild<W14.SdtContentCheckBox>()!
            .GetFirstChild<W14.Checked>()!;
        isChecked.Val = value ? W14.OnOffValues.One : W14.OnOffValues.Zero;
        var text = sdt.Descendants<Text>().First();
        text.Text = value ? "☒" : "☐";
    }

    static void SetFullDate(Body body, string tag, DateTime value)
    {
        var sdt = FindSdt(body, tag);
        sdt.SdtProperties!.GetFirstChild<SdtContentDate>()!.FullDate = value;
    }

    static void ClearFullDate(Body body, string tag)
    {
        var sdt = FindSdt(body, tag);
        sdt.SdtProperties!.GetFirstChild<SdtContentDate>()!.FullDate = null;
    }
}
