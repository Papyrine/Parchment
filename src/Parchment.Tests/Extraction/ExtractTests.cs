using System.Globalization;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

public class ExtractTests
{
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
        store.RegisterDocxTemplate<EditableFieldTests.EditableQuote>("extract-nested", template);

        using var stream = new MemoryStream();
        await store.Render(
            "extract-nested",
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

    static async Task<MemoryStream> RenderOrder(
        EditableFieldTests.EditableOrder model,
        [CallerMemberName] string name = "")
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
        store.RegisterDocxTemplate<EditableFieldTests.EditableOrder>(name, template);

        var stream = new MemoryStream();
        await store.Render(name, model, stream);
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
}
