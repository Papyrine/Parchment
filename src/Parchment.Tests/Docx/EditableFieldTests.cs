using DocumentFormat.OpenXml.Validation;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using SdtLock = DocumentFormat.OpenXml.Wordprocessing.Lock;
// ReSharper disable PartialTypeWithSinglePart

public partial class EditableFieldTests
{
    static string ScenarioPath(string scenarioName) =>
        Path.Combine(
            ProjectFiles.ProjectDirectory,
            "Scenarios",
            scenarioName);

    #region EditableFieldsModel
    [ParchmentBindable]
    public partial class OrderForm
    {
        public required string Number;

        [EditableField]
        public required string PurchaseOrder { get; set; }

        [EditableField]
        public bool Approved { get; set; }

        [EditableField]
        public OrderStatus Status { get; set; }

        [EditableField(MultiLine = true)]
        public string? Notes { get; set; }
    }

    public enum OrderStatus
    {
        Draft,
        Submitted,
        Accepted
    }
    #endregion

    [Test]
    public async Task Scenario()
    {
        #region EditableFieldsUsage
        var templatePath = Path.Combine(ScenarioPath("editable-fields"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<OrderForm>(templatePath);

        var model = new OrderForm
        {
            Number = "ORD-2026-042",
            PurchaseOrder = "PO-77041",
            Approved = true,
            Status = OrderStatus.Submitted,
            Notes = null
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        #endregion

        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("editable-fields"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }

    [Test]
    public async Task ScenarioRoundTrip()
    {
        var templatePath = Path.Combine(ScenarioPath("editable-fields"), "input.docx");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<OrderForm>(templatePath);

        var model = new OrderForm
        {
            Number = "ORD-2026-042",
            PurchaseOrder = "PO-77041",
            Approved = false,
            Status = OrderStatus.Draft,
            Notes = null
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        // Simulate the user filling in the form in Word.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            SetSdtText(FindSdt(body, "PurchaseOrder"), "PO-91000");
            SetSdtChecked(FindSdt(body, "Approved"), true);
            SetSdtText(FindSdt(body, "Status"), "Accepted");
            SetSdtText(FindSdt(body, "Notes"), "Ship to the loading dock.");
            doc.Save();
        }

        stream.Position = 0;

        #region EditableFieldsExtract
        var result = ParchmentExtractor.Extract<OrderForm>(stream);

        result.ApplyTo(model);
        #endregion

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(model.PurchaseOrder).IsEqualTo("PO-91000");
        await Assert.That(model.Approved).IsTrue();
        await Assert.That(model.Status).IsEqualTo(OrderStatus.Accepted);
        await Assert.That(model.Notes).IsEqualTo("Ship to the loading dock.");
    }

    static void SetSdtText(SdtRun sdt, string text)
    {
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

    static void SetSdtChecked(SdtRun sdt, bool value)
    {
        sdt.SdtProperties!
            .GetFirstChild<W14.SdtContentCheckBox>()!
            .GetFirstChild<W14.Checked>()!
            .Val = value ? W14.OnOffValues.One : W14.OnOffValues.Zero;
        sdt.Descendants<Text>().First().Text = value ? "☒" : "☐";
    }

    public enum QuoteStatus
    {
        Draft,
        Submitted,
        Accepted
    }

    [ParchmentBindable]
    public partial class EditableOrder
    {
        public required string Number;
        public required List<string> Tags;
        public required bool IncludeNotes;

        [EditableField]
        public required string PurchaseOrder { get; set; }

        [EditableField]
        public bool Approved { get; set; }

        [EditableField]
        public Date Delivery { get; set; }

        [EditableField(DateFormat = "yyyy-MM-dd HH:mm")]
        public DateTime DispatchedAt { get; set; }

        [EditableField]
        public DateTimeOffset SignedAt { get; set; }

        [EditableField]
        public Time PickupTime { get; set; }

        [EditableField]
        public QuoteStatus Status { get; set; }

        [EditableField]
        public decimal Discount { get; set; }

        [EditableField]
        public string? Notes { get; set; }

        [EditableField(MultiLine = true)]
        public string? Instructions { get; set; }
    }

    public static EditableOrder NewOrder() =>
        new()
        {
            Number = "ORD-100",
            Tags = ["a", "b"],
            IncludeNotes = true,
            PurchaseOrder = "PO-2026-17",
            Approved = true,
            Delivery = new(2026, 7, 6),
            DispatchedAt = new(2026, 7, 6, 14, 30, 0),
            SignedAt = new(2026, 7, 6, 9, 0, 0, TimeSpan.FromHours(10)),
            PickupTime = new(16, 45, 0),
            Status = QuoteStatus.Submitted,
            Discount = 10m,
            Notes = null,
            Instructions = "Line one\nLine two"
        };

    [Test]
    public async Task RenderAllKinds()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Order {{ Number }}

            PO: {{ PurchaseOrder }}

            Approved: {{ Approved }}

            Delivery: {{ Delivery }}

            Status: {{ Status }}

            Discount: {{ Discount }}

            Notes: {{ Notes }}

            {{ Instructions }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task SdtStructurePerKind()
    {
        using var stream = await Render(
            """
            PO: {{ PurchaseOrder }}

            {{ Approved }}

            {{ Delivery }}

            {{ Status }}

            {{ Discount }}

            {{ Notes }}

            {{ Instructions }}
            """,
            NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.Descendants<SdtRun>().Count()).IsEqualTo(7);

        var po = FindSdt(body, "PurchaseOrder");
        await Assert.That(po.SdtProperties!.GetFirstChild<SdtAlias>()!.Val!.Value).IsEqualTo("PurchaseOrder");
        await Assert.That(po.SdtProperties.GetFirstChild<SdtLock>()!.Val!.Value).IsEqualTo(LockingValues.SdtLocked);
        await Assert.That(po.SdtProperties.GetFirstChild<SdtContentText>()).IsNotNull();
        await Assert.That(po.InnerText).IsEqualTo("PO-2026-17");

        var approved = FindSdt(body, "Approved");
        var checkbox = approved.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>();
        await Assert.That(checkbox).IsNotNull();
        await Assert.That(checkbox!.GetFirstChild<W14.Checked>()!.Val!.Value).IsEqualTo(W14.OnOffValues.One);
        await Assert.That(approved.InnerText).IsEqualTo("☒");

        var delivery = FindSdt(body, "Delivery");
        var date = delivery.SdtProperties!.GetFirstChild<SdtContentDate>();
        await Assert.That(date).IsNotNull();
        await Assert.That(date!.FullDate!.Value).IsEqualTo(new(2026, 7, 6));
        await Assert.That(date.GetFirstChild<DateFormat>()!.Val!.Value).IsEqualTo("yyyy-MM-dd");
        await Assert.That(delivery.InnerText).IsEqualTo("2026-07-06");

        var status = FindSdt(body, "Status");
        var dropDown = status.SdtProperties!.GetFirstChild<SdtContentDropDownList>();
        await Assert.That(dropDown).IsNotNull();
        var items = dropDown!.Elements<ListItem>().Select(_ => _.Value!.Value!).ToList();
        await Assert.That(items).IsEquivalentTo(["Draft", "Submitted", "Accepted"]);
        await Assert.That(status.InnerText).IsEqualTo("Submitted");

        var discount = FindSdt(body, "Discount");
        await Assert.That(discount.InnerText).IsEqualTo("10");

        // Null nullable value renders as a placeholder.
        var notes = FindSdt(body, "Notes");
        await Assert.That(notes.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();

        var instructions = FindSdt(body, "Instructions");
        var text = instructions.SdtProperties!.GetFirstChild<SdtContentText>();
        await Assert.That(text!.MultiLine!.Value).IsTrue();
        await Assert.That(instructions.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(instructions.InnerText).IsEqualTo("Line oneLine two");
    }

    [ParchmentBindable]
    public partial class PlaceholderKindsModel
    {
        [EditableField]
        public QuoteStatus? Choice { get; set; }

        [EditableField]
        public Date? When { get; set; }

        [EditableField]
        public string? Text { get; set; }
    }

    [Test]
    public async Task PlaceholderTextMatchesControlKind()
    {
        using var stream = await RenderModel(
            """
            {{ Choice }}

            {{ When }}

            {{ Text }}
            """,
            new PlaceholderKindsModel());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // Each unset control shows the placeholder for its own kind — a dropdown reading "Enter text"
        // looks like a broken text field.
        await Assert.That(FindSdt(body, "Choice").InnerText).IsEqualTo("Choose an item");
        await Assert.That(FindSdt(body, "When").InnerText).IsEqualTo("Enter a date");
        await Assert.That(FindSdt(body, "Text").InnerText).IsEqualTo("Enter text");
    }

    [ParchmentBindable]
    public partial class StyledPlaceholderModel
    {
        [EditableField]
        public string? Text { get; set; }

        [EditableField(MultiLine = true)]
        public string? Lines { get; set; }

        [EditableField]
        public decimal? Amount { get; set; }

        [EditableField]
        public Date? When { get; set; }

        [EditableField]
        public DateTimeOffset? SignedAt { get; set; }

        [EditableField]
        public Time? PickupTime { get; set; }

        [EditableField]
        public QuoteStatus? Choice { get; set; }

        [Html]
        [EditableField]
        public string? Body { get; set; }
    }

    [Test]
    public async Task PlaceholderInCharacterStyledRunHasOneRunStyle()
    {
        string[] inline = ["Text", "Lines", "Amount", "When", "SignedAt", "PickupTime", "Choice"];

        // Every token sits in a run carrying a character style, the shape Word produces when the
        // template author styles the token.
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Text }}

            {{ Lines }}

            {{ Amount }}

            {{ When }}

            {{ SignedAt }}

            {{ PickupTime }}

            {{ Choice }}

            {{ Body }}
            """);
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            foreach (var run in doc.MainDocumentPart!.Document!.Body!.Descendants<Run>())
            {
                run.RunProperties ??= new();
                run.RunProperties.RunStyle = new()
                {
                    Val = "Strong"
                };
            }

            doc.Save();
        }

        template.Position = 0;
        var store = new TemplateStore();
        store.RegisterDocxTemplate<StyledPlaceholderModel>(template);

        using var stream = new MemoryStream();
        await store.Render(new StyledPlaceholderModel(), stream);
        stream.Position = 0;

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            foreach (var tag in inline)
            {
                var sdt = FindSdt(body, tag);
                await Assert.That(sdt.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();

                // The placeholder run is styled as a placeholder only...
                var styles = sdt.Descendants<Run>().Single().RunProperties!.Elements<RunStyle>().ToList();
                await Assert.That(styles.Count).IsEqualTo(1);
                await Assert.That(styles[0].Val!.Value).IsEqualTo("PlaceholderText");

                // ...while the control keeps the site's character style for the text typed over it.
                var controlStyles = sdt.SdtProperties.GetFirstChild<RunProperties>()!.Elements<RunStyle>().ToList();
                await Assert.That(controlStyles.Count).IsEqualTo(1);
                await Assert.That(controlStyles[0].Val!.Value).IsEqualTo("Strong");
            }

            var block = FindSdtBlock(body, "Body");
            await Assert.That(block.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();
            foreach (var properties in block.Descendants<RunProperties>())
            {
                await Assert.That(properties.Elements<RunStyle>().Count()).IsEqualTo(1);
            }

            var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
            var errors = validator.Validate(doc)
                .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
                .ToList();
            await Assert.That(errors).IsEmpty();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<StyledPlaceholderModel>(stream);
        foreach (var tag in inline.Append("Body"))
        {
            await Assert.That(result.Fields.Single(_ => _.Path == tag).State).IsEqualTo(FieldState.Empty);
        }
    }

    public enum ReviewStage
    {
        NotYetStarted,

        [Display(Name = "Under review")]
        InProgress
    }

    [ParchmentBindable]
    public partial class ReviewModel
    {
        [EditableField]
        public ReviewStage Stage { get; set; }
    }

    [Test]
    public async Task DropDownShowsFriendlyEnumLabelsButRoundTripsTheMemberName()
    {
        using var stream = await RenderModel("{{ Stage }}", new ReviewModel { Stage = ReviewStage.InProgress });

        using var doc = WordprocessingDocument.Open(stream, false);
        var sdt = FindSdt(doc.MainDocumentPart!.Document!.Body!, "Stage");
        var items = sdt.SdtProperties!.GetFirstChild<SdtContentDropDownList>()!.Elements<ListItem>().ToList();

        // DisplayText is friendly (humanized, or the [Display] name); Value stays the member name.
        await Assert.That(items.Select(_ => _.DisplayText!.Value!)).IsEquivalentTo(["Not yet started", "Under review"]);
        await Assert.That(items.Select(_ => _.Value!.Value!)).IsEquivalentTo(["NotYetStarted", "InProgress"]);
        // The visible text is the selected item's friendly label.
        await Assert.That(sdt.InnerText).IsEqualTo("Under review");

        // ...and it round-trips back to the enum value via the list item's Value.
        stream.Position = 0;
        var result = ParchmentExtractor.Extract<ReviewModel>(stream);
        await Assert.That(result.Fields.Single(_ => _.Path == "Stage").Value).IsEqualTo(ReviewStage.InProgress);
    }

    [Test]
    public async Task DropDownUserPickingAFriendlyLabelExtractsToTheEnum()
    {
        // Render one value, then simulate the user choosing a different option in Word — which replaces
        // the control text with that item's friendly label ("Under review"), NOT the member name.
        // Extraction must resolve the label back through the w:listItem, so the read value differs from
        // what was rendered — proving it isn't just echoing the rendered member.
        using var stream = await RenderModel("{{ Stage }}", new ReviewModel { Stage = ReviewStage.NotYetStarted });

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var sdt = FindSdt(doc.MainDocumentPart!.Document!.Body!, "Stage");
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            content.RemoveAllChildren();
            content.AppendChild(new Run(new Text("Under review")));
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<ReviewModel>(stream);
        await Assert.That(result.Fields.Single(_ => _.Path == "Stage").Value).IsEqualTo(ReviewStage.InProgress);
    }

    public enum AliasedLevel
    {
        Low = 1,
        Minor = 1,

        [Display(Name = "Top")]
        High = 2,
        VeryHigh = 2,
        Other = 3
    }

    public enum SameLabelChoice
    {
        [Display(Name = "Same")]
        First,

        [Display(Name = "Same")]
        Second,
        Third
    }

    [ParchmentBindable]
    public partial class AliasedModel
    {
        [EditableField]
        public AliasedLevel Level { get; set; }

        [EditableField]
        public AliasedLevel? Optional { get; set; }

        [EditableField]
        public SameLabelChoice Choice { get; set; }
    }

    static List<(string Text, string Value)> DropDownItems(SdtRun sdt) =>
        sdt.SdtProperties!.GetFirstChild<SdtContentDropDownList>()!
            .Elements<ListItem>()
            .Select(_ => (_.DisplayText!.Value!, _.Value!.Value!))
            .ToList();

    static void PickDropDownText(Stream stream, string tag, string text)
    {
        using var doc = WordprocessingDocument.Open(stream, true);
        var sdt = FindSdt(doc.MainDocumentPart!.Document!.Body!, tag);
        var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
        content.RemoveAllChildren();
        content.AppendChild(new Run(new Text(text)));
        doc.Save();
    }

    [Test]
    public async Task DropDownGivesEachNameOfASharedEnumValueItsOwnLabel()
    {
        // Minor is an alias of Low and VeryHigh of High; the model holds the alias both times.
        using var stream = await RenderModel(
            "{{ Level }} {{ Optional }}",
            new AliasedModel { Level = AliasedLevel.Minor, Optional = AliasedLevel.VeryHigh });

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var level = FindSdt(body, "Level");
            var items = DropDownItems(level);
            await Assert.That(items.Select(_ => _.Value)).IsEquivalentTo(["Low", "Minor", "High", "VeryHigh", "Other"]);
            await Assert.That(items.Select(_ => _.Text)).IsEquivalentTo(["Low", "Minor", "Top", "Very high", "Other"]);

            // The selection is the first declared name of the value.
            await Assert.That(level.InnerText).IsEqualTo("Low");
            await Assert.That(FindSdt(body, "Optional").InnerText).IsEqualTo("Top");
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<AliasedModel>(stream);
        await Assert.That(result.Fields.Single(_ => _.Path == "Level").Value).IsEqualTo(AliasedLevel.Low);
        await Assert.That(result.Fields.Single(_ => _.Path == "Optional").Value).IsEqualTo(AliasedLevel.High);

        // Picking an alias entry extracts its value.
        stream.Position = 0;
        PickDropDownText(stream, "Level", "Very high");
        stream.Position = 0;
        result = ParchmentExtractor.Extract<AliasedModel>(stream);
        await Assert.That(result.Fields.Single(_ => _.Path == "Level").Value).IsEqualTo(AliasedLevel.High);
    }

    [Test]
    public async Task DropDownKeepsMembersWithTheSameLabelDistinguishable()
    {
        using var stream = await RenderModel("{{ Choice }}", new AliasedModel { Choice = SameLabelChoice.Second });

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var choice = FindSdt(doc.MainDocumentPart!.Document!.Body!, "Choice");
            await Assert.That(DropDownItems(choice).Select(_ => _.Text)).IsEquivalentTo(["Same", "Same (Second)", "Third"]);
            await Assert.That(choice.InnerText).IsEqualTo("Same (Second)");
        }

        // Second used to read back as First, the first item showing "Same".
        stream.Position = 0;
        var result = ParchmentExtractor.Extract<AliasedModel>(stream);
        await Assert.That(result.Fields.Single(_ => _.Path == "Choice").Value).IsEqualTo(SameLabelChoice.Second);

        stream.Position = 0;
        PickDropDownText(stream, "Choice", "Same");
        stream.Position = 0;
        result = ParchmentExtractor.Extract<AliasedModel>(stream);
        await Assert.That(result.Fields.Single(_ => _.Path == "Choice").Value).IsEqualTo(SameLabelChoice.First);
    }

    [Test]
    public async Task TemporalKindsRenderWithCorrectControls()
    {
        using var stream = await Render(
            """
            {{ DispatchedAt }}

            {{ SignedAt }}

            {{ PickupTime }}
            """,
            NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // DateTime keeps the native date picker; its time-of-day survives in w:fullDate even
        // though the (custom) format also shows it in the visible text.
        var dispatched = FindSdt(body, "DispatchedAt");
        var date = dispatched.SdtProperties!.GetFirstChild<SdtContentDate>();
        await Assert.That(date).IsNotNull();
        await Assert.That(date!.FullDate!.Value).IsEqualTo(new(2026, 7, 6, 14, 30, 0));
        await Assert.That(dispatched.InnerText).IsEqualTo("2026-07-06 14:30");

        // DateTimeOffset has no offset-aware picker and w:fullDate can't hold an offset, so it
        // renders as plain text carrying a round-trippable ISO value (offset included).
        var signed = FindSdt(body, "SignedAt");
        await Assert.That(signed.SdtProperties!.GetFirstChild<SdtContentText>()).IsNotNull();
        await Assert.That(signed.SdtProperties.GetFirstChild<SdtContentDate>()).IsNull();
        await Assert.That(signed.InnerText).IsEqualTo("2026-07-06T09:00:00+10:00");

        // TimeOnly has no picker either — plain text.
        var pickup = FindSdt(body, "PickupTime");
        await Assert.That(pickup.SdtProperties!.GetFirstChild<SdtContentText>()).IsNotNull();
        await Assert.That(pickup.SdtProperties.GetFirstChild<SdtContentDate>()).IsNull();
        await Assert.That(pickup.InnerText).IsEqualTo("16:45:00");
    }

    [Test]
    public async Task EveryFieldGetsAPairedEditableRange()
    {
        using var stream = await Render(
            """
            {{ PurchaseOrder }}

            {{ Approved }}
            """,
            NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var starts = body.Descendants<PermStart>().ToList();
        var ends = body.Descendants<PermEnd>().ToList();
        await Assert.That(starts.Count).IsEqualTo(2);
        await Assert.That(ends.Count).IsEqualTo(2);

        var startIds = starts.Select(_ => _.Id!.Value).ToList();
        await Assert.That(startIds.Distinct().Count()).IsEqualTo(2);
        await Assert.That(ends.Select(_ => _.Id!.Value)).IsEquivalentTo(startIds);
        foreach (var start in starts)
        {
            await Assert.That(start.EditorGroup!.Value).IsEqualTo(RangePermissionEditingGroupValues.Everyone);
        }

        // Deterministic ids also key the sdt controls.
        var sdtIds = body.Descendants<SdtRun>()
            .Select(_ => _.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value)
            .ToList();
        await Assert.That(sdtIds).IsEquivalentTo(startIds);
    }

    [Test]
    public async Task ProtectionAppliedWhenEditable()
    {
        using var stream = await Render("{{ PurchaseOrder }}", NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var protection = doc.MainDocumentPart!.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentProtection>();
        await Assert.That(protection).IsNotNull();
        await Assert.That(protection!.Edit!.Value).IsEqualTo(DocumentProtectionValues.ReadOnly);
        await Assert.That(protection.Enforcement!.Value).IsTrue();
    }

    [Test]
    public async Task ProtectionSkippedWhenModeNone()
    {
        using var template = DocxTemplateBuilder.Build("{{ PurchaseOrder }}");
        var store = new TemplateStore();

        #region ProtectionModeNone
        store.RegisterDocxTemplate<EditableOrder>(template, ProtectionMode.None);
        #endregion

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var protection = doc.MainDocumentPart!.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentProtection>();
        await Assert.That(protection).IsNull();

        // The field itself still renders as a tagged control.
        await Assert.That(doc.MainDocumentPart.Document!.Body!.Descendants<SdtRun>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task ProtectionNotAppliedWithoutEditableFields()
    {
        using var template = DocxTemplateBuilder.Build("Invoice {{ Number }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var protection = doc.MainDocumentPart!.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentProtection>();
        await Assert.That(protection).IsNull();
    }

    const string legacyPassword = "w:cryptProviderType=\"rsaAES\" w:cryptAlgorithmClass=\"hash\" w:cryptAlgorithmType=\"typeAny\" w:cryptAlgorithmSid=\"14\" w:cryptSpinCount=\"100000\" w:hash=\"q2FzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhcw==\" w:salt=\"c2FsdHNhbHRzYWx0c2FsdA==\"";
    const string modernPassword = "w:algorithmName=\"SHA-512\" w:hashValue=\"q2FzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhc2RmYXNkZmFzZGZhcw==\" w:saltValue=\"c2FsdHNhbHRzYWx0c2FsdA==\" w:spinCount=\"100000\"";

    [Test]
    [Arguments("forms", legacyPassword)]
    [Arguments("readOnly", legacyPassword)]
    [Arguments("comments", legacyPassword)]
    [Arguments("trackedChanges", legacyPassword)]
    [Arguments("none", legacyPassword)]
    [Arguments("forms", modernPassword)]
    [Arguments("readOnly", modernPassword)]
    public async Task ProtectionKeepsTheTemplatePassword(string edit, string password)
    {
        // Elements on both sides, so the kept element has to land back in its schema slot.
        using var stream = await RenderProtected(
            $"<w:zoom w:percent=\"100\"/><w:documentProtection w:edit=\"{edit}\" w:formatting=\"1\" w:enforcement=\"0\" {password}/><w:defaultTabStop w:val=\"720\"/>");

        using var doc = WordprocessingDocument.Open(stream, false);
        var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        var protection = settings.Elements<DocumentProtection>().Single();

        // Whatever the author chose, the render needs enforced read-only for the editable ranges.
        await Assert.That(protection.Edit!.Value).IsEqualTo(DocumentProtectionValues.ReadOnly);
        await Assert.That(protection.Enforcement!.Value).IsTrue();

        var kept = protection.GetAttributes()
            .Where(_ => _.LocalName is not ("edit" or "enforcement"))
            .Select(_ => $"w:{_.LocalName}=\"{_.Value}\"");
        await Assert.That(string.Join(' ', kept)).IsEqualTo($"w:formatting=\"1\" {password}");

        var names = settings.ChildElements.Select(_ => _.LocalName).Take(3);
        await Assert.That(string.Join(',', names)).IsEqualTo("zoom,documentProtection,defaultTabStop");
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc).ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task ProtectionWithoutAPasswordIsReplaced()
    {
        using var stream = await RenderProtected(
            "<w:zoom w:percent=\"100\"/><w:documentProtection w:edit=\"forms\" w:formatting=\"1\" w:enforcement=\"1\"/><w:defaultTabStop w:val=\"720\"/>");

        using var doc = WordprocessingDocument.Open(stream, false);
        var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        var protection = settings.Elements<DocumentProtection>().Single();
        await Assert.That(protection.Edit!.Value).IsEqualTo(DocumentProtectionValues.ReadOnly);
        await Assert.That(protection.Enforcement!.Value).IsTrue();
        await Assert.That(protection.GetAttributes().Count).IsEqualTo(2);
    }

    [Test]
    public async Task ProtectionModeNoneLeavesTheTemplateProtectionAlone()
    {
        using var stream = await RenderProtected(
            $"<w:documentProtection w:edit=\"forms\" w:enforcement=\"1\" {legacyPassword}/>",
            ProtectionMode.None);

        using var doc = WordprocessingDocument.Open(stream, false);
        var protection = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!.Elements<DocumentProtection>().Single();
        await Assert.That(protection.Edit!.Value).IsEqualTo(DocumentProtectionValues.Forms);
        await Assert.That(protection.Hash).IsNotNull();
    }

    // Renders an editable field from a template whose settings already hold the given children.
    static async Task<MemoryStream> RenderProtected(string settingsChildren, ProtectionMode mode = ProtectionMode.WhenEditable)
    {
        using var template = DocxTemplateBuilder.Build("{{ PurchaseOrder }}");
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var main = doc.MainDocumentPart!;
            var part = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
            part.Settings = new($"<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{settingsChildren}</w:settings>");
            doc.Save();
        }

        template.Position = 0;

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template, mode);

        var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task HtmlEditableSeedsListDefinitions()
    {
        using var stream = await RenderModel(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = "<p>plain</p>"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart?.Numbering;

        // The value has no list, yet protection has locked numbering.xml — so a definition must be
        // seeded, or Word would disable bullets inside the editable block.
        await Assert.That(numbering).IsNotNull();
        var formats = numbering!.Elements<AbstractNum>()
            .Select(_ => _.Elements<Level>().First(level => level.LevelIndex?.Value == 0).NumberingFormat!.Val!.Value)
            .ToList();
        await Assert.That(formats).Contains(NumberFormatValues.Bullet);
        await Assert.That(formats).Contains(NumberFormatValues.Decimal);
    }

    [Test]
    public async Task PlainEditableDoesNotSeedListDefinitions()
    {
        using var stream = await Render("{{ PurchaseOrder }}", NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        // No rich-text field, so nothing to seed — the plain-text control can't hold a list anyway.
        await Assert.That(doc.MainDocumentPart!.NumberingDefinitionsPart).IsNull();
    }

    [Test]
    public async Task HtmlEditableSeedsStyleDefinitions()
    {
        // The test fixture ships a full styles part, but a real code-built template (like the bid export)
        // ships none — strip it so the seed has something to do.
        using var stream = await RenderStylesStripped(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = "<p>plain</p>"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart?.Styles;

        // Protection has locked styles.xml, so the heading styles must be seeded or Word would grey out
        // the style gallery inside the editable block.
        await Assert.That(styles).IsNotNull();
        var styleIds = styles!.Elements<Style>().Select(_ => _.StyleId?.Value).ToList();
        await Assert.That(styleIds).Contains("Heading1");
        await Assert.That(styleIds).Contains("ListParagraph");
    }

    [Test]
    public async Task PlainEditableDoesNotSeedStyleDefinitions()
    {
        using var stream = await RenderStylesStripped("{{ PurchaseOrder }}", NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        // No rich-text field, so nothing to seed — the stripped styles part stays absent.
        await Assert.That(doc.MainDocumentPart!.StyleDefinitionsPart).IsNull();
    }

    // Renders after removing the fixture's styles part, so style seeding is observable (the fixture
    // otherwise ships every style the seed would add, making it a no-op).
    static async Task<MemoryStream> RenderStylesStripped<T>(string templateContent, T model)
    {
        using var template = DocxTemplateBuilder.Build(templateContent);
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var stylesPart = doc.MainDocumentPart!.StyleDefinitionsPart;
            if (stylesPart != null)
            {
                doc.MainDocumentPart.DeletePart(stylesPart);
            }

            doc.Save();
        }

        template.Position = 0;

        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);

        var stream = new MemoryStream();
        await store.Render(model!, stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task NonSoloTokenPreservesSurroundingText()
    {
        using var stream = await Render("PO: {{ PurchaseOrder }} (required)", NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).IsEqualTo("PO: PO-2026-17 (required)");
        await Assert.That(FindSdt(body, "PurchaseOrder").InnerText).IsEqualTo("PO-2026-17");
    }

    [Test]
    public async Task TwoEditableFieldsInOneParagraph()
    {
        using var stream = await Render("{{ PurchaseOrder }} / {{ Discount }}", NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).IsEqualTo("PO-2026-17 / 10");
        await Assert.That(body.Descendants<SdtRun>().Count()).IsEqualTo(2);
        await Assert.That(body.Descendants<PermStart>().Count()).IsEqualTo(2);
        await Assert.That(body.Descendants<PermEnd>().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task TokenStraddlingRunsSplicesCorrectly()
    {
        // Word splits text into multiple runs when formatting changes mid-token. The splicer
        // must reassemble the token span across runs and keep the surrounding halves.
        using var template = BuildStraddledTemplate();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).IsEqualTo("PO: PO-2026-17 end");
        await Assert.That(FindSdt(body, "PurchaseOrder").InnerText).IsEqualTo("PO-2026-17");
    }

    public static IEnumerable<Func<(string Name, Func<OpenXmlElement[], OpenXmlElement> Wrap)>> InlineWrappers()
    {
        yield return () => ("hyperlink", _ => new Hyperlink(_)
        {
            Anchor = "top"
        });
        yield return () => ("sdt", _ => new SdtRun(
            new SdtProperties(
                new Tag
                {
                    Val = "author"
                }),
            new SdtContentRun(_)));
        yield return () => ("ins", _ => new InsertedRun(_)
        {
            Id = "901",
            Author = "author"
        });
        yield return () => ("moveTo", _ => new MoveToRun(_)
        {
            Id = "902",
            Author = "author"
        });
        yield return () => ("fldSimple", _ => new SimpleField(_)
        {
            Instruction = " AUTHOR "
        });
        yield return () => ("dir", _ => new BidirectionalEmbedding(_)
        {
            Val = DirectionValues.Ltr
        });
        yield return () => ("bdo", _ => new BidirectionalOverride(_)
        {
            Val = DirectionValues.Ltr
        });
    }

    [Test]
    [MethodDataSource(nameof(InlineWrappers))]
    public async Task TokenInsideInlineWrapperStaysInPlace(string name, Func<OpenXmlElement[], OpenXmlElement> wrap)
    {
        // The token's run is a child of the wrapper, not of the paragraph. The control has to
        // land between the wrapper's own "link: " and " here", not after the wrapper's text.
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear ")),
            wrap([new Run(Preserved("link: {{ PurchaseOrder }} here"))]),
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo($"Dear <{name}>link: [PO-2026-17] here</{name}> end");
    }

    [Test]
    [MethodDataSource(nameof(InlineWrappers))]
    public async Task TokenFillingInlineWrapperStaysInside(string name, Func<OpenXmlElement[], OpenXmlElement> wrap)
    {
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear ")),
            wrap([new Run(Preserved("{{ PurchaseOrder }}"))]),
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo($"Dear <{name}>[PO-2026-17]</{name}> end");
    }

    [Test]
    [MethodDataSource(nameof(InlineWrappers))]
    public async Task TokenSplitAcrossRunsInsideInlineWrapper(string name, Func<OpenXmlElement[], OpenXmlElement> wrap)
    {
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear ")),
            wrap(
            [
                new Run(Preserved("link: {{ Purch")),
                new Run(
                    new RunProperties(new Bold()),
                    Preserved("aseOrder }} here"))
            ]),
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo($"Dear <{name}>link: [PO-2026-17] here</{name}> end");
    }

    [Test]
    [MethodDataSource(nameof(InlineWrappers))]
    public async Task TokenStartingInsideInlineWrapperLandsInside(string name, Func<OpenXmlElement[], OpenXmlElement> wrap)
    {
        // A token straddling the wrapper's edge: the control goes where the token starts.
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear ")),
            wrap([new Run(Preserved("link: {{ Purch"))]),
            new Run(Preserved("aseOrder }} end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo($"Dear <{name}>link: [PO-2026-17]</{name}> end");
    }

    [Test]
    [MethodDataSource(nameof(InlineWrappers))]
    public async Task TokenStartingBeforeInlineWrapperLandsBefore(string name, Func<OpenXmlElement[], OpenXmlElement> wrap)
    {
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear {{ Purch")),
            wrap([new Run(Preserved("aseOrder }} here"))]),
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo($"Dear [PO-2026-17]<{name}> here</{name}> end");
    }

    [Test]
    public async Task TwoTokensInsideOneInlineWrapper()
    {
        using var template = BuildParagraphTemplate(
            new Run(Preserved("{{ Discount }} ")),
            new Hyperlink(
                new Run(Preserved("a {{ PurchaseOrder }} b {{ Delivery }} c")))
            {
                Anchor = "top"
            },
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo("[10] <hyperlink>a [PO-2026-17] b [2026-07-06] c</hyperlink> end");
    }

    [Test]
    public async Task TokenInsideNestedInlineWrappersStaysInPlace()
    {
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear ")),
            new Hyperlink(
                new InsertedRun(
                    new Run(Preserved("link: {{ PurchaseOrder }} here")))
                {
                    Id = "901",
                    Author = "author"
                })
            {
                Anchor = "top"
            },
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo("Dear <hyperlink><ins>link: [PO-2026-17] here</ins></hyperlink> end");
    }

    public static IEnumerable<Func<(string Kind, OpenXmlElement Marker)>> ValueControlKinds()
    {
        yield return () => ("plain-text", new SdtContentText());
        yield return () => ("plain-text", new SdtContentText
        {
            MultiLine = true
        });
        yield return () => ("date picker", new SdtContentDate());
        yield return () => ("drop-down list", new SdtContentDropDownList());
        yield return () => ("combo box", new SdtContentComboBox());
        yield return () => ("checkbox", new W14.SdtContentCheckBox());
        yield return () => ("picture", new SdtContentPicture());
    }

    static SdtRun ValueControl(OpenXmlElement marker, string text) =>
        new(
            new SdtProperties(
                new Tag
                {
                    Val = "author"
                },
                marker),
            new SdtContentRun(new Run(Preserved(text))));

    [Test]
    [MethodDataSource(nameof(ValueControlKinds))]
    public async Task TokenInsideValueControlIsRejected(string kind, OpenXmlElement marker)
    {
        // Word refuses to open a document with a control nested in one that holds a single
        // value, so the field cannot be placed where the token is.
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear ")),
            ValueControl(marker, "po: {{ PurchaseOrder }} here"),
            new Run(Preserved(" end")));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EditableOrder>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains($"sits inside a {kind} content control");
    }

    [Test]
    public async Task TokenInsideBlockLevelValueControlIsRejected()
    {
        using var template = BuildBodyTemplate(
            new SdtBlock(
                new SdtProperties(new SdtContentText()),
                new SdtContentBlock(
                    new Paragraph(
                        new Run(Preserved("po: {{ PurchaseOrder }}"))))));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EditableOrder>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("sits inside a plain-text content control");
    }

    [Test]
    public async Task TokenStartingBeforeValueControlIsAllowed()
    {
        // The field goes where the token starts, which is outside the control.
        using var template = BuildParagraphTemplate(
            new Run(Preserved("Dear {{ Purch")),
            ValueControl(new SdtContentText(), "aseOrder }} here"),
            new Run(Preserved(" end")));

        await Assert.That(await RenderWrapperLayout(template))
            .IsEqualTo("Dear [PO-2026-17]<sdt> here</sdt> end");
    }

    [Test]
    public async Task ReadOnlyTokensInsideValueControlAreAllowed()
    {
        // Neither becomes a content control: Number is not editable, and an editable member
        // inside a loop renders as read-only text.
        using var template = BuildBodyTemplate(
            new Paragraph(
                ValueControl(new SdtContentText(), "{{ Number }}")),
            new Paragraph(new Run(Preserved("{% for tag in Tags %}"))),
            new Paragraph(
                ValueControl(new SdtContentText(), "{{ PurchaseOrder }}")),
            new Paragraph(new Run(Preserved("{% endfor %}"))));

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).IsEqualTo("ORD-100PO-2026-17PO-2026-17");
        await Assert.That(body.Descendants<PermStart>().Any()).IsFalse();
    }

    [Test]
    public async Task CollectionItemTokenInsideValueControlIsRejected()
    {
        using var template = BuildBodyTemplate(
            new Paragraph(new Run(Preserved("{% for budget in Budgets %}"))),
            new Paragraph(
                ValueControl(new SdtContentText(), "{{ budget.Year }}")),
            new Paragraph(new Run(Preserved("{% endfor %}"))));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EditableCollectionTests.BudgetPlan>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("sits inside a plain-text content control");
    }

    [Test]
    public async Task TabAfterTokenInSameRunIsKept()
    {
        // The token's run is left holding only the tab once the token text is removed; it is
        // not an empty run and must survive, after the control.
        using var template = BuildParagraphTemplate(
            new Run(
                Preserved("{{ PurchaseOrder }}"),
                new TabChar(),
                Preserved("after")));

        await Assert.That(await RenderLayout(template)).IsEqualTo("[PO-2026-17]<tab>after");
    }

    [Test]
    public async Task BreakAfterTokenEndingParagraphIsKept()
    {
        using var template = BuildParagraphTemplate(
            new Run(
                Preserved("before"),
                new TabChar(),
                Preserved("{{ PurchaseOrder }}"),
                new Break()));

        await Assert.That(await RenderLayout(template)).IsEqualTo("before<tab>[PO-2026-17]<br>");
    }

    [Test]
    public async Task NonTextContentAroundMidTextTokenStaysOnItsSide()
    {
        using var template = BuildParagraphTemplate(
            new Run(
                new RunProperties(new Bold()),
                new TabChar(),
                Preserved("x {{ PurchaseOrder }}"),
                new Break(),
                new TabChar(),
                Preserved(" y {{ Discount }}"),
                new TabChar()));

        await Assert.That(await RenderLayout(template))
            .IsEqualTo("<tab>x [PO-2026-17]<br><tab> y [10]<tab>");
    }

    [Test]
    public async Task RunEmptiedByTokenIsRemoved()
    {
        using var template = BuildParagraphTemplate(
            new Run(Preserved("PO: ")),
            new Run(
                new RunProperties(new Bold()),
                Preserved("{{ PurchaseOrder }}")),
            new Run(Preserved(" end")));
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("PO: PO-2026-17 end");
        await Assert.That(paragraph.Elements<Run>().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task NestedPathBindsAndTags()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Ref {{ Reference }}

            Email: {{ Customer.ContactEmail }}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableQuote>(template);

        var model = new EditableQuote
        {
            Reference = "Q-7",
            Customer = new()
            {
                ContactEmail = "ada@example.com"
            }
        };
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(FindSdt(body, "Customer.ContactEmail").InnerText).IsEqualTo("ada@example.com");
    }

    [Test]
    public async Task NullParentRendersEmptyControls()
    {
        using var stream = await RenderModel(nullParentTemplate, new EditableAccount());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // The getter short-circuits to null through the null Profile. A checkbox has no placeholder
        // state, so it renders unchecked; every other kind shows its placeholder.
        var active = FindSdt(body, "Profile.Active");
        var checkBox = active.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!;
        await Assert.That(checkBox.GetFirstChild<W14.Checked>()!.Val!.Value).IsEqualTo(W14.OnOffValues.Zero);
        await Assert.That(active.InnerText).IsEqualTo("☐");
        await Assert.That(active.SdtProperties.GetFirstChild<ShowingPlaceholder>()).IsNull();

        await Assert.That(FindSdt(body, "Profile.Seats").InnerText).IsEqualTo("Enter text");
        await Assert.That(FindSdt(body, "Profile.Joined").InnerText).IsEqualTo("Enter a date");
        await Assert.That(FindSdt(body, "Profile.Status").InnerText).IsEqualTo("Choose an item");
        await Assert.That(FindSdt(body, "Profile.Name").InnerText).IsEqualTo("Enter text");
    }

    [Test]
    public async Task NullParentCheckboxRoundTrip()
    {
        using var stream = await RenderModel(nullParentTemplate, new EditableAccount());

        var result = ParchmentExtractor.Extract<EditableAccount>(stream);
        await Assert.That(result.AllExtracted).IsTrue();

        // The unchecked box extracts as a real false (a checkbox is never Empty), so applying it
        // needs a reachable parent - the same reachability rule as any other extracted value.
        var active = result.Fields.First(_ => _.Path == "Profile.Active");
        await Assert.That(active.State).IsEqualTo(FieldState.Extracted);
        await Assert.That((bool)active.Value!).IsFalse();
        await Assert.That(result.Fields.First(_ => _.Path == "Profile.Seats").State).IsEqualTo(FieldState.Empty);

        var unreachable = new EditableAccount();
        var exception = await Assert.That(() => result.ApplyTo(unreachable))
            .Throws<ParchmentExtractionException>();
        await Assert.That(exception!.Message).Contains("Profile.Active");
        await Assert.That(unreachable.Profile).IsNull();

        var reachable = new EditableAccount
        {
            Profile = new()
            {
                Active = true,
                Seats = 3
            }
        };
        result.ApplyTo(reachable);
        await Assert.That(reachable.Profile.Active).IsFalse();
        await Assert.That(reachable.Profile.Seats).IsEqualTo(3);
    }

    [Test]
    public async Task EditableInsideChosenIfBranchRenders()
    {
        var template =
            """
            {% if IncludeNotes %}

            Notes: {{ Notes }}

            {% else %}

            No notes on {{ Number }}.

            {% endif %}
            """;

        var model = NewOrder();
        model.Notes = "call before delivery";
        using var chosen = await Render(template, model);
        using (var doc = WordprocessingDocument.Open(chosen, false))
        {
            await Assert.That(FindSdt(doc.MainDocumentPart!.Document!.Body!, "Notes").InnerText)
                .IsEqualTo("call before delivery");
        }

        var without = NewOrder();
        without.IncludeNotes = false;
        using var eliminated = await Render(template, without);
        using (var doc = WordprocessingDocument.Open(eliminated, false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            await Assert.That(body.Descendants<SdtRun>().Any()).IsFalse();
            await Assert.That(body.InnerText).IsEqualTo("No notes on ORD-100.");
        }
    }

    [Test]
    public async Task EditableInsideLoopRendersPlainText()
    {
        using var stream = await Render(
            """
            {% for tag in Tags %}

            {{ PurchaseOrder }}

            {% endfor %}
            """,
            NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.Descendants<SdtRun>().Any()).IsFalse();
        await Assert.That(body.InnerText).IsEqualTo("PO-2026-17PO-2026-17");
    }

    [Test]
    public async Task HeaderOccurrenceRendersPlainText()
    {
        // Word does not reliably honor editable-range exceptions outside the body, so header /
        // footer occurrences of an editable member render as plain read-only substitutions.
        // This also means the body occurrence stays the single tagged control for extraction.
        using var template = BuildTemplateWithHeader(
            bodyText: "{{ PurchaseOrder }}",
            headerText: "PO: {{ PurchaseOrder }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Count()).IsEqualTo(1);

        var header = doc.MainDocumentPart.HeaderParts.Single().Header!;
        await Assert.That(header.Descendants<SdtRun>().Any()).IsFalse();
        await Assert.That(header.InnerText).IsEqualTo("PO: PO-2026-17");
    }

    // The plain-access and once-only rules exist for the control. A loop body and a header render
    // the member as read-only text through Fluid, so a filter is applied there and a repeat is
    // only a mirror of the one body control.
    [Test]
    public async Task ReadOnlyLoopOccurrenceTakesAFilterBesideTheBodyControl()
    {
        using var stream = await Render(
            """
            {{ PurchaseOrder }}

            {% for tag in Tags %}

            {{ PurchaseOrder | downcase }}

            {% endfor %}
            """,
            NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.Descendants<SdtRun>().Count()).IsEqualTo(1);
        await Assert.That(body.InnerText).IsEqualTo("PO-2026-17po-2026-17po-2026-17");
    }

    [Test]
    public async Task HeaderOccurrenceTakesAFilter()
    {
        using var template = BuildTemplateWithHeader(
            bodyText: "{{ PurchaseOrder }}",
            headerText: "PO: {{ PurchaseOrder | downcase }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Count()).IsEqualTo(1);
        await Assert.That(doc.MainDocumentPart.HeaderParts.Single().Header!.InnerText).IsEqualTo("PO: po-2026-17");
    }

    [Test]
    public async Task OutputValidates()
    {
        using var stream = await Render(
            """
            PO: {{ PurchaseOrder }} ({{ Discount }})

            {{ Approved }} {{ Delivery }} {{ Status }}

            {{ Notes }}
            """,
            NewOrder());

        using var doc = WordprocessingDocument.Open(stream, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    // A template can hold content controls and editable ranges of its own. Their ids share one
    // space with the ones a render adds, across the body, headers and footers, and the render used
    // to number from 1 regardless.
    [Test]
    public async Task IdsStartAboveTheTemplatesOwn()
    {
        using var template = DocxTemplateBuilder.Build("PO: {{ PurchaseOrder }}\n\n{{ Approved }} {{ Status }}");
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var main = doc.MainDocumentPart!;
            main.Document!.Body!.InsertAt(
                new Paragraph(
                    new PermStart
                    {
                        Id = 1,
                        EditorGroup = RangePermissionEditingGroupValues.Everyone
                    },
                    new SdtRun(
                        new SdtProperties(
                            new Tag
                            {
                                Val = "authorControl"
                            },
                            new SdtId
                            {
                                Val = 2
                            }),
                        new SdtContentRun(new Run(new Text("author's own control")))),
                    new PermEnd
                    {
                        Id = 1
                    }),
                0);

            // Token-free, so a render never reads it: the ids have to be found at registration.
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new(
                new Paragraph(
                    new PermStart
                    {
                        Id = 6,
                        EditorGroup = RangePermissionEditingGroupValues.Everyone
                    },
                    new SdtRun(
                        new SdtProperties(
                            new SdtId
                            {
                                Val = 4
                            }),
                        new SdtContentRun(new Run(new Text("header control")))),
                    new PermEnd
                    {
                        Id = 6
                    }));
            // Word writes ids anywhere in the signed range; a negative one is not a floor.
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new(
                new Paragraph(
                    new SdtRun(
                        new SdtProperties(
                            new SdtId
                            {
                                Val = -12345
                            }),
                        new SdtContentRun(new Run(new Text("footer control"))))));
            doc.Save();
        }

        template.Position = 0;
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var first = new MemoryStream();
        await store.Render(NewOrder(), first);
        using var second = new MemoryStream();
        await store.Render(NewOrder(), second);
        await Assert.That(first.ToArray().SequenceEqual(second.ToArray())).IsTrue();

        first.Position = 0;
        using var rendered =WordprocessingDocument.Open(first, false);
        var renderedMain = rendered.MainDocumentPart!;
        var roots = new List<OpenXmlCompositeElement>
        {
            renderedMain.Document!.Body!
        };
        roots.AddRange(renderedMain.HeaderParts.Select(_ => _.Header!));
        roots.AddRange(renderedMain.FooterParts.Select(_ => _.Footer!));

        var sdtIds = roots
            .SelectMany(_ => _.Descendants<SdtId>())
            .Select(_ => _.Val!.Value)
            .Order()
            .ToList();
        var permIds = roots
            .SelectMany(_ => _.Descendants<PermStart>())
            .Select(_ => _.Id!.Value)
            .Order()
            .ToList();
        await Assert.That(sdtIds).IsEquivalentTo([-12345, 2, 4, 7, 8, 9]);
        await Assert.That(permIds).IsEquivalentTo([1, 6, 7, 8, 9]);

        foreach (var start in roots.SelectMany(_ => _.Descendants<PermStart>()))
        {
            var ends = roots
                .SelectMany(_ => _.Descendants<PermEnd>())
                .Count(_ => _.Id!.Value == start.Id!.Value);
            await Assert.That(ends).IsEqualTo(1);
        }

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(rendered)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task FilterOnEditableTokenIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("{{ PurchaseOrder | upcase }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EditableOrder>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("plain member-access");
    }

    [Test]
    public async Task DuplicateEditableTokenIsRejected()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ PurchaseOrder }}

            Again: {{ PurchaseOrder }}
            """);
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EditableOrder>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }


    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;


    [ParchmentBindable]
    public partial class MixedStructuralModel
    {
        [Html]
        public string Body { get; set; } = "";

        [EditableField]
        public string PurchaseOrder { get; set; } = "";
    }

    [Test]
    public async Task EditableSharingParagraphWithStructuralTokenIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("{{ Body }} and {{ PurchaseOrder }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<MixedStructuralModel>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("own paragraph");
    }

    #region EditableRichTextModel
    [ParchmentBindable]
    public partial class EditableArticle
    {
        public required string Title;

        [Html]
        [EditableField]
        public required string Body { get; set; }
    }
    #endregion


    [Test]
    public async Task HtmlEditableRendersUnlockedEditableBlock()
    {
        using var stream = await RenderModel(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = "<p>Hello <strong>world</strong></p>"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var sdt = FindSdtBlock(body, "Body");
        await Assert.That(sdt.SdtProperties!.GetFirstChild<SdtAlias>()!.Val!.Value).IsEqualTo("Body");
        // No sdtLocked: Word refuses multi-paragraph list formatting on a locked control when the
        // selection includes its first or last paragraph, and multi-paragraph formatted content is
        // the whole point of a rich-text field.
        await Assert.That(sdt.SdtProperties.GetFirstChild<SdtLock>()).IsNull();
        // Rich content: the formatting survives into the editable block.
        await Assert.That(sdt.Descendants<Bold>().Any()).IsTrue();
        await Assert.That(sdt.InnerText).IsEqualTo("Hello world");
        // Wrapped in a single editable perm range.
        await Assert.That(body.Descendants<PermStart>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<PermEnd>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task HtmlEditableAloneInTableCellUsesWholeCellPermRange()
    {
        using var stream = await RenderTableModel(
            "Body:",
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = "<p>one</p><p>two</p><p>three</p>"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // The control fills the cell; the editable exception is the whole cell (row-level
        // colFirst/colLast), not a block range inside it — a tight range excludes the
        // end-of-cell marker, which makes Word silently refuse list formatting on the
        // cell's first/last paragraphs.
        var sdt = FindSdtBlock(body, "Body");
        var cell = (TableCell)sdt.Parent!;
        await Assert.That(cell.Elements<Paragraph>().Any()).IsFalse();
        await Assert.That(cell.Descendants<PermStart>().Any()).IsFalse();
        await Assert.That(cell.Descendants<PermEnd>().Any()).IsFalse();

        var row = (TableRow)cell.Parent!;
        var start = row.Elements<PermStart>().Single();
        await Assert.That(start.ColumnFirst!.Value).IsEqualTo(1);
        await Assert.That(start.ColumnLast!.Value).IsEqualTo(1);
        await Assert.That(start.EditorGroup!.Value).IsEqualTo(RangePermissionEditingGroupValues.Everyone);
        await Assert.That(row.Elements<PermEnd>().Single().Id!.Value).IsEqualTo(start.Id!.Value);
        // permStart precedes the cells; permEnd follows them.
        await Assert.That(row.ChildElements[0]).IsEqualTo(start);
        await Assert.That(row.ChildElements[^1]).IsTypeOf<PermEnd>();

        var sdtId = sdt.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value;
        await Assert.That(start.Id!.Value).IsEqualTo(sdtId);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        var field = result.Fields.Single(_ => _.Path == "Body");
        await Assert.That(field.Value).IsEqualTo("<p>one</p><p>two</p><p>three</p>");
    }

    [Test]
    public async Task HtmlEditableHoldingTableAloneInTableCellIsFollowedByParagraph()
    {
        const string html = "<p>intro</p><table><tr><td>x</td></tr></table>";
        using var stream = await RenderTableModel(
            "Body:",
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = html
            });

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            // The control's content ends with the table, and Word wants the cell to end with a
            // paragraph. It goes after the control and not into it, where it would be read back
            // as part of the value.
            var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
            await Assert.That(sdt.SdtContentBlock!.LastChild).IsTypeOf<Table>();
            var cell = (TableCell)sdt.Parent!;
            await Assert.That(cell.ChildElements.Count).IsEqualTo(2);
            await Assert.That(cell.LastChild).IsTypeOf<Paragraph>();
            await Assert.That(cell.LastChild!.InnerText).IsEqualTo("");
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        var field = result.Fields.Single(_ => _.Path == "Body");
        await Assert.That(field.State).IsEqualTo(FieldState.Extracted);
        // What the same value reads back as from a control in the body, where nothing is added.
        await Assert.That(field.Value).IsEqualTo(await ExtractHtml(html));
    }

    [ParchmentBindable]
    public partial class RichTextShapes
    {
        [Html]
        [EditableField]
        public required string BodyHosted { get; set; }

        [Html]
        [EditableField]
        public required string CellHosted { get; set; }
    }

    [Test]
    public async Task RenderHtmlEditableShapes()
    {
        // Snapshots the two rich-text shapes end to end: a body-hosted control (tight perm wrap,
        // no sdtLocked) and a sole-content-of-cell control (row-level whole-cell perm range) —
        // the markup Word verification showed is required for list formatting to work on every
        // paragraph of a protected rich-text field.
        using var template = DocxTemplateBuilder.BuildWithTable(
            "Cell hosted:",
            "{{ CellHosted }}",
            bodyParagraphText: "{{ BodyHosted }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<RichTextShapes>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new RichTextShapes
            {
                BodyHosted = "<p>body intro</p><ul><li>body first</li><li>body second</li></ul>",
                CellHosted = "<p>cell intro</p><ul><li>cell first</li><li>cell second</li></ul>"
            },
            stream);

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task HtmlEditableSharingTableCellKeepsBlockPermRange()
    {
        using var stream = await RenderTableModel(
            "Body:",
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = "<p>content</p>"
            },
            extraValueCellParagraph: true);

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        // A sibling paragraph shares the cell, so a whole-cell range would expose it to
        // editing — the tight block range around the control is kept instead.
        var sdt = FindSdtBlock(body, "Body");
        var cell = (TableCell)sdt.Parent!;
        await Assert.That(cell.Elements<PermStart>().Single().ColumnFirst).IsNull();
        await Assert.That(cell.Elements<PermEnd>().Count()).IsEqualTo(1);
        var row = (TableRow)cell.Parent!;
        await Assert.That(row.Elements<PermStart>().Any()).IsFalse();
        await Assert.That(row.Elements<PermEnd>().Any()).IsFalse();
    }

    static async Task<MemoryStream> RenderTableModel<T>(
        string labelText,
        string valueCellText,
        T model,
        bool extraValueCellParagraph = false)
    {
        using var template = DocxTemplateBuilder.BuildWithTable(labelText, valueCellText, extraValueCellParagraph);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);

        var stream = new MemoryStream();
        await store.Render(model!, stream);
        stream.Position = 0;
        return stream;
    }

    [Test]
    public async Task HtmlEditableOutputValidates()
    {
        using var stream = await RenderModel(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = "<p>Hello <strong>world</strong></p><ul><li>a</li><li>b</li></ul>"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task HtmlEditableRoundTripsThroughExtraction()
    {
        var model = new EditableArticle
        {
            Title = "T",
            Body = "<p>Hello <strong>world</strong> and <em>more</em></p><ul><li>one</li><li>two</li></ul>"
        };
        using var stream = await RenderModel("{{ Body }}", model);

        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        result.ApplyTo(model);

        await Assert.That(result.AllExtracted).IsTrue();
        await Verify(model.Body)
            .Snapshot("<p>Hello <strong>world</strong> and <em>more</em></p><ul><li>one</li><li>two</li></ul>");
    }

    [Test]
    public Task HtmlEditableBulletListThenOrderedListRoundTrips() =>
        AssertHtmlRoundTrips("<ul><li>a</li></ul><ol><li>b</li></ol>");

    [Test]
    public Task HtmlEditableOrderedListThenBulletListRoundTrips() =>
        AssertHtmlRoundTrips("<ol><li>a</li><li>b</li></ol><ul><li>c</li></ul>");

    [Test]
    public Task HtmlEditableOrderedListNestedInBulletItemRoundTrips() =>
        AssertHtmlRoundTrips("<ul><li>a<ol><li>b</li><li>c</li></ol></li><li>d</li></ul>");

    [Test]
    public Task HtmlEditableNestedListsOfDifferentKindsRoundTrip() =>
        AssertHtmlRoundTrips("<ul><li>a<ul><li>b</li></ul><ol><li>c</li></ol></li></ul><ol><li>d</li></ol>");

    [Test]
    public async Task HtmlEditableAdjacentListsOfTheSameKindExtractAsOne()
    {
        // The numbering instance is not a list boundary: Word hands out a new w:numId for bullets
        // a user sees as one list, so only a change of kind splits.
        var extracted = await ExtractHtml("<ul><li>a</li></ul><ul><li>b</li></ul>");

        await Assert.That(extracted).IsEqualTo("<ul><li>a</li><li>b</li></ul>");
    }

    // Unedited content is a fixed point: the bold a heading is drawn with and the underline a link
    // is drawn with are implied by the element, not formatting to report on top of it.
    [Test]
    public Task HtmlEditableSupportedSubsetRoundTrips() =>
        AssertHtmlRoundTrips(
            "<h1>Title</h1>" +
            "<h2>Sub <em>stressed</em> <u>under</u> <a href=\"https://example.com/h\">link</a></h2>" +
            "<p>Hello <strong>bold</strong> <em>ital</em> <u>under</u> <s>strike</s> x<sup>2</sup> H<sub>2</sub>O<br>next</p>" +
            "<p>see <a href=\"https://example.com/a?x=1&amp;y=2\">link</a> and " +
            "<a href=\"https://example.com/b\"><strong>bold</strong> link</a> and <a href=\"#local\">local</a></p>" +
            "<ul><li>a <a href=\"https://example.com/c\">c</a><ol><li><strong>b</strong></li></ol></li></ul>" +
            "<h3>Three</h3><h4>Four</h4><h5>Five</h5><h6>Six</h6>");

    [Test]
    public async Task HtmlEditableImpliedFormattingIsNotReported()
    {
        var extracted = await ExtractHtml(
            "<h1><strong>Title</strong></h1><p><a href=\"https://example.com/a\"><u>link</u></a></p>");

        await Assert.That(extracted).IsEqualTo("<h1>Title</h1><p><a href=\"https://example.com/a\">link</a></p>");
    }

    [Test]
    public async Task HtmlEditableLinkWithUnderlineRemovedRoundTrips()
    {
        // Ctrl+U over a link the render wrote drops its direct w:u.
        var extracted = await ExtractHtml(
            "<p><a href=\"https://example.com/a\">link</a></p>",
            content =>
            {
                foreach (var underline in content.Descendants<Underline>().ToList())
                {
                    underline.Remove();
                }
            });

        await Assert.That(extracted).IsEqualTo("<p><a href=\"https://example.com/a\" style=\"text-decoration:none\">link</a></p>");
        await AssertHtmlRoundTrips(extracted!);
    }

    [Test]
    public async Task HtmlEditableLinkWithUnderlineRemovedFromPartRoundTrips()
    {
        var extracted = await ExtractHtml(
            "<p><a href=\"https://example.com/a\"><strong>one</strong> two</a></p>",
            content =>
                content.Descendants<Hyperlink>().Single()
                    .Elements<Run>().First()
                    .RunProperties!.Underline = new()
                {
                    Val = UnderlineValues.None
                });

        await Assert.That(extracted).IsEqualTo("<p><a href=\"https://example.com/a\" style=\"text-decoration:none\"><strong>one</strong><u> two</u></a></p>");
        await AssertHtmlRoundTrips(extracted!);
    }

    [Test]
    public async Task HtmlEditableLinkInsertedInWordIsUnderlinedByItsStyle()
    {
        // Word draws its own links through the Hyperlink character style, with no direct w:u.
        var extracted = await ExtractHtml(
            "<p><a href=\"https://example.com/a\">link</a></p>",
            content =>
                content.Descendants<Hyperlink>().Single()
                    .Elements<Run>().Single()
                    .RunProperties = new(
                    new RunStyle
                    {
                        Val = "Hyperlink"
                    }));

        await Assert.That(extracted).IsEqualTo("<p><a href=\"https://example.com/a\">link</a></p>");
    }

    [Test]
    public async Task HtmlEditableHeadingRunWithBoldSwitchedOffRoundTrips()
    {
        var extracted = await ExtractHtml(
            "<h1>one <em>two</em></h1>",
            content =>
                content.Descendants<Run>().Last()
                    .RunProperties!.Bold = new()
                {
                    Val = false
                });

        await Assert.That(extracted).IsEqualTo("<h1>one <span style=\"font-weight:normal\"><em>two</em></span></h1>");
        await AssertHtmlRoundTrips(extracted!);
    }

    [Test]
    public async Task HtmlEditableParagraphGivenHeadingStyleInWordIsAPlainHeading()
    {
        // The style gallery sets w:pStyle and leaves the runs alone, so they carry no w:b.
        var extracted = await ExtractHtml(
            "<p>Title <strong>bold</strong></p>",
            content =>
                content.Descendants<Paragraph>().Single()
                    .ParagraphProperties = new(
                    new ParagraphStyleId
                    {
                        Val = "Heading2"
                    }));

        await Assert.That(extracted).IsEqualTo("<h2>Title bold</h2>");
    }

    [Test]
    public Task HtmlEditableOrderedListStartRoundTrips() =>
        AssertHtmlRoundTrips("<ol start=\"5\"><li>five</li><li>six</li></ol>");

    [Test]
    public Task HtmlEditableOrderedListStartingAtZeroRoundTrips() =>
        AssertHtmlRoundTrips("<ol start=\"0\"><li>zero</li><li>one</li></ol>");

    [Test]
    public Task HtmlEditableNestedOrderedListStartRoundTrips() =>
        AssertHtmlRoundTrips("<ol start=\"3\"><li>a<ol start=\"7\"><li>b</li><li>c</li></ol></li><li>d</li></ol>");

    [Test]
    public Task HtmlEditableOrderedListStartNestedInBulletItemRoundTrips() =>
        AssertHtmlRoundTrips("<ul><li>a<ol start=\"4\"><li>b</li></ol></li><li>c</li></ul>");

    [Test]
    public Task HtmlEditableOrderedListStartAfterBulletListRoundTrips() =>
        AssertHtmlRoundTrips("<ul><li>a</li></ul><ol start=\"2\"><li>b</li></ol>");

    [Test]
    public Task HtmlEditableOrderedListsEitherSideOfParagraphKeepTheirOwnStart() =>
        AssertHtmlRoundTrips("<ol start=\"5\"><li>a</li></ol><p>mid</p><ol><li>b</li></ol><p>end</p><ol start=\"9\"><li>c</li></ol>");

    [Test]
    public async Task HtmlEditableOrderedListStartOfOneExtractsWithoutStart()
    {
        // The render writes a start override of 1 on every ordered list; that is html's default,
        // not an attribute.
        var extracted = await ExtractHtml("<ol start=\"1\"><li>a</li></ol>");

        await Assert.That(extracted).IsEqualTo("<ol><li>a</li></ol>");
    }

    [Test]
    public async Task HtmlEditableOrderedListStartReadsDefinitionLevelStart()
    {
        // A list numbered in Word can carry its first number on the definition's level, with no
        // override on the instance.
        var model = new EditableArticle
        {
            Title = "T",
            Body = "<ol start=\"5\"><li>a</li><li>b</li></ol>"
        };
        using var stream = await RenderModel("{{ Body }}", model);

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var main = doc.MainDocumentPart!;
            var numberingId = FindSdtBlock(main.Document!.Body!, "Body")
                .Descendants<NumberingId>()
                .First()
                .Val!.Value;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            var instance = numbering.Elements<NumberingInstance>().Single(_ => _.NumberID!.Value == numberingId);
            instance.RemoveAllChildren<LevelOverride>();
            numbering
                .Elements<AbstractNum>()
                .Single(_ => _.AbstractNumberId!.Value == instance.AbstractNumId!.Val!.Value)
                .Elements<Level>()
                .Single(_ => _.LevelIndex!.Value == 0)
                .StartNumberingValue = new() { Val = 8 };
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        result.ApplyTo(model);

        await Assert.That(model.Body).IsEqualTo("<ol start=\"8\"><li>a</li><li>b</li></ol>");
    }

    [Test]
    [Arguments("mailto:a@b.co?subject=Hi%20there")]
    [Arguments("tel:+61-2-5550-1234")]
    [Arguments("https://example.com")]
    [Arguments("https://example.com/a%20b?q=%C3%A9#frag")]
    public async Task HtmlEditableLinkKeepsItsHref(string href)
    {
        // The relationship target is read back as written. Uri.ToString() unescapes it, turning
        // %20 into a space, and adds the slash an authority-only url did not have.
        var extracted = await ExtractHtml($"<p><a href=\"{href}\">text</a></p>");
        await Assert.That(extracted).Contains($"<a href=\"{href}\">");
    }

    static async Task AssertHtmlRoundTrips(string html)
    {
        var extracted = await ExtractHtml(html);

        await Assert.That(extracted).IsEqualTo(html);
    }

    static async Task<string?> ExtractHtml(string html, Action<OpenXmlElement>? edit = null)
    {
        var model = new EditableArticle
        {
            Title = "T",
            Body = html
        };
        using var stream = await RenderModel("{{ Body }}", model);

        if (edit != null)
        {
            using (var doc = WordprocessingDocument.Open(stream, true))
            {
                var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
                edit(sdt.ChildElements.First(_ => _.LocalName == "sdtContent"));
                doc.Save();
            }

            stream.Position = 0;
        }

        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        model.Body = "";
        result.ApplyTo(model);
        return model.Body;
    }

    [Test]
    public async Task HtmlEditableEditedContentRoundTrips()
    {
        var model = new EditableArticle
        {
            Title = "T",
            Body = "<p>original</p>"
        };
        using var stream = await RenderModel("{{ Body }}", model);

        // Simulate the user editing the rich block in Word.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
            sdt.SdtProperties!.RemoveAllChildren<ShowingPlaceholder>();
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            content.RemoveAllChildren();
            content.AppendChild(
                new Paragraph(
                    new Run(
                        new RunProperties(new Bold()),
                        new Text("edited"))));
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        result.ApplyTo(model);

        await Assert.That(model.Body).IsEqualTo("<p><strong>edited</strong></p>");
    }

    const string wordXmlns =
        """xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" """;

    const string revision = """w:id="901" w:author="u" w:date="2026-01-01T00:00:00Z" """;

    static string WordRun(string text) =>
        $"""<w:r><w:t xml:space="preserve">{text}</w:t></w:r>""";

    // Extraction reads a document with every tracked change accepted. The link run carries the
    // Hyperlink style, as one inserted in Word does: without it the link reads as not underlined.
    [Test]
    public async Task HtmlEditableKeepsTrackedInsertionsAndDropsDeletions()
    {
        var extracted = await ExtractEditedHtml(
            $"""
             <w:p>{WordRun("keep ")}<w:ins {revision}>{WordRun("typed")}</w:ins><w:del {revision}><w:r><w:delText> gone</w:delText></w:r></w:del></w:p>
             <w:p><w:pPr><w:rPr><w:ins {revision}/></w:rPr></w:pPr><w:ins {revision}>{WordRun("new paragraph")}</w:ins></w:p>
             <w:p><w:pPr><w:rPr><w:del {revision}/></w:rPr></w:pPr><w:del {revision}><w:r><w:delText>gone paragraph</w:delText></w:r></w:del></w:p>
             <w:p><w:hyperlink w:anchor="bm"><w:ins {revision}><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>link</w:t></w:r></w:ins></w:hyperlink></w:p>
             """);

        await Assert.That(extracted).IsEqualTo("""<p>keep typed</p><p>new paragraph</p><p><a href="#bm">link</a></p>""");
    }

    [Test]
    public async Task HtmlEditableReadsMovedTextAtItsDestination()
    {
        var extracted = await ExtractEditedHtml(
            $"""
             <w:p><w:pPr><w:rPr><w:moveFrom {revision}/></w:rPr></w:pPr><w:moveFrom {revision}>{WordRun("moved")}</w:moveFrom></w:p>
             <w:p>{WordRun("stay")}</w:p>
             <w:p><w:moveTo {revision}>{WordRun("moved")}</w:moveTo></w:p>
             """);

        await Assert.That(extracted).IsEqualTo("<p>stay</p><p>moved</p>");
    }

    [Test]
    public async Task HtmlEditableReadsThroughInlineWrappers()
    {
        var extracted = await ExtractEditedHtml(
            $"""
             <w:p>{WordRun("on ")}<w:fldSimple w:instr=" DATE ">{WordRun("6/07/2026")}</w:fldSimple></w:p>
             <w:p><w:smartTag w:uri="urn:x" w:element="place">{WordRun("Paris")}</w:smartTag>{WordRun(" trip")}</w:p>
             <w:p><w:dir w:val="rtl">{WordRun("dir ")}</w:dir><w:bdo w:val="ltr">{WordRun("bdo ")}</w:bdo><w:customXml w:uri="urn:x" w:element="e">{WordRun("custom")}</w:customXml></w:p>
             <w:p>{WordRun("a ")}<w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>inner</w:t></w:r></w:sdtContent></w:sdt></w:p>
             """);

        await Assert.That(extracted).IsEqualTo("<p>on 6/07/2026</p><p>Paris trip</p><p>dir bdo custom</p><p>a <strong>inner</strong></p>");
    }

    [Test]
    public async Task HtmlEditableReadsThroughBlockWrappers()
    {
        var extracted = await ExtractEditedHtml(
            $"""
             <w:p>{WordRun("a")}</w:p>
             <w:sdt><w:sdtPr><w:tag w:val="inner"/></w:sdtPr><w:sdtContent>
               <w:p>{WordRun("inner block")}</w:p>
               <w:tbl><w:tr><w:tc><w:p>{WordRun("cell")}</w:p></w:tc></w:tr><w:tr><w:trPr><w:del {revision}/></w:trPr><w:tc><w:p>{WordRun("deleted row")}</w:p></w:tc></w:tr></w:tbl>
             </w:sdtContent></w:sdt>
             <w:customXml w:uri="urn:x" w:element="e"><w:p>{WordRun("custom block")}</w:p></w:customXml>
             """);

        await Assert.That(extracted).IsEqualTo("<p>a</p><p>inner block</p><p>cell</p><p>custom block</p>");
    }

    // Neither has an HTML form in the supported subset, so both degrade to their text. A text box
    // is stored twice (the shape and its VML fallback) and is read once.
    [Test]
    public async Task HtmlEditableKeepsTheTextOfMathAndTextBoxes()
    {
        var boxed = $"<w:txbxContent><w:p>{WordRun("boxed")}</w:p></w:txbxContent>";
        var extracted = await ExtractEditedHtml(
            $"""
             <w:p>{WordRun("math ")}<m:oMath><m:r><m:t>x&lt;1</m:t></m:r></m:oMath></w:p>
             <w:p>{WordRun("before ")}<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:wsp><wps:txbx>{boxed}</wps:txbx></wps:wsp></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox>{boxed}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
             """);

        await Assert.That(extracted).IsEqualTo("<p>math x&lt;1</p><p>before boxed</p>");
    }

    static async Task<string?> ExtractEditedHtml(string contentXml)
    {
        var model = new EditableArticle
        {
            Title = "T",
            Body = "<p>original</p>"
        };
        using var stream = await RenderModel("{{ Body }}", model);

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
            sdt.SdtProperties!.RemoveAllChildren<ShowingPlaceholder>();
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            sdt.ReplaceChild(new SdtContentBlock($"<w:sdtContent {wordXmlns}>{contentXml}</w:sdtContent>"), content);
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        result.ApplyTo(model);
        return model.Body;
    }

    const string fragmentUrl = "https://example.com/docs/page#section-two";

    [Test]
    public async Task HtmlEditableLinkFragmentSplitByWordIsRejoined()
    {
        // On any save Word moves a link's fragment out of the relationship target and into
        // w:anchor on the w:hyperlink.
        using var stream = await RenderLink(fragmentUrl);
        RewriteLink(stream, "https://example.com/docs/page", _ => _.Anchor = "section-two");

        var extracted = ExtractBody(stream);

        await Assert.That(extracted).Contains($"<a href=\"{fragmentUrl}\">");

        // The extracted html renders back to the link the document started with.
        using var again = await RenderLink(extracted!, wrap: false);
        await Assert.That(LinkTargets(again)).IsEquivalentTo([fragmentUrl]);
    }

    [Test]
    public async Task HtmlEditableLinkFragmentInDocLocationIsRejoined()
    {
        using var stream = await RenderLink(fragmentUrl);
        RewriteLink(stream, "https://example.com/docs/page", _ => _.DocLocation = "section-two");

        await Assert.That(ExtractBody(stream)).Contains($"<a href=\"{fragmentUrl}\">");
    }

    [Test]
    public async Task HtmlEditableLinkAnchorWinsOverDocLocation()
    {
        using var stream = await RenderLink(fragmentUrl);
        RewriteLink(
            stream,
            "https://example.com/docs/page",
            _ =>
            {
                _.Anchor = "section-two";
                _.DocLocation = "elsewhere";
            });

        await Assert.That(ExtractBody(stream)).Contains($"<a href=\"{fragmentUrl}\">");
    }

    [Test]
    public async Task HtmlEditableLinkFragmentLeftInTargetExtractsOnce()
    {
        // The shape Parchment writes, before Word has saved the document.
        using var stream = await RenderLink(fragmentUrl);

        await Assert.That(ExtractBody(stream)).Contains($"<a href=\"{fragmentUrl}\">");
    }

    [Test]
    public async Task HtmlEditableBookmarkLinkExtractsAsFragmentOnly()
    {
        using var stream = await RenderLink(fragmentUrl);
        RewriteLink(stream, null, _ => _.Anchor = "section-two");

        await Assert.That(ExtractBody(stream)).Contains("<a href=\"#section-two\">");
    }

    [Test]
    public async Task HtmlEditableLinkKeepsPercentEscapesAndEscapesTheAttribute()
    {
        using var stream = await RenderLink(fragmentUrl);
        RewriteLink(stream, "https://example.com/q?a=%22x%22&b=Hi%20there", _ => _.Anchor = "a<b");

        await Assert.That(ExtractBody(stream))
            .Contains("<a href=\"https://example.com/q?a=%22x%22&amp;b=Hi%20there#a&lt;b\">");
    }

    static Task<MemoryStream> RenderLink(string value, bool wrap = true)
    {
        var body = value;
        if (wrap)
        {
            body = $"<p>See <a href=\"{value}\">section two</a> for more</p>";
        }

        return RenderModel(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = body
            });
    }

    /// <summary>
    /// Points the rendered link at <paramref name="target"/> through a new relationship (none when
    /// null) and lets the caller set the attributes Word keeps on the <c>w:hyperlink</c> itself.
    /// </summary>
    static void RewriteLink(MemoryStream stream, string? target, Action<Hyperlink> edit)
    {
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var hyperlink = FindSdtBlock(mainPart.Document!.Body!, "Body")
                .Descendants<Hyperlink>()
                .Single();
            mainPart.DeleteReferenceRelationship(hyperlink.Id!.Value!);
            hyperlink.Id = null;
            if (target != null)
            {
                mainPart.AddHyperlinkRelationship(new(target, UriKind.Absolute), true, "rIdWord1");
                hyperlink.Id = "rIdWord1";
            }

            edit(hyperlink);
            doc.Save();
        }

        stream.Position = 0;
    }

    static string? ExtractBody(MemoryStream stream)
    {
        var field = ParchmentExtractor.Extract<EditableArticle>(stream)
            .Fields.Single(_ => _.Path == "Body");
        stream.Position = 0;
        return field.Value as string;
    }

    static List<string> LinkTargets(MemoryStream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.HyperlinkRelationships
            .Select(_ => _.Uri.OriginalString)
            .ToList();
    }

    [Test]
    public async Task HtmlEditableTabRoundTrips()
    {
        // The bare character would fold to a space in html, so a tab is the span OpenXmlHtml
        // renders as w:tab — and what it renders has to extract as the same span again.
        const string html = "<p>left<span style=\"white-space:pre\">&#9;</span>right</p>";
        var model = new EditableArticle
        {
            Title = "T",
            Body = html
        };
        using var stream = await RenderModel("{{ Body }}", model);

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
            await Assert.That(sdt.Descendants<TabChar>().Count()).IsEqualTo(1);
            await Assert.That(sdt.InnerText).IsEqualTo("leftright");
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        model.Body = "";
        result.ApplyTo(model);

        await Assert.That(model.Body).IsEqualTo(html);
    }

    [Test]
    public async Task HtmlEditableReadsTheRunContentWordWritesAsElements()
    {
        var model = new EditableArticle
        {
            Title = "T",
            Body = "<p>original</p>"
        };
        using var stream = await RenderModel("{{ Body }}", model);

        // Word writes each of these as an element of its own rather than as text in w:t.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            content.RemoveAllChildren();
            content.AppendChild(
                new Paragraph(
                    new Run(
                        new Text("a"),
                        new TabChar(),
                        new Text("b")),
                    new Run(
                        new RunProperties(new Bold()),
                        new TabChar()),
                    new Run(
                        new Text("co"),
                        new NoBreakHyphen(),
                        new Text("op"),
                        new SoftHyphen(),
                        new Text("s"),
                        new CarriageReturn(),
                        new SymbolChar
                        {
                            Font = "Arial",
                            Char = "2022"
                        },
                        new SymbolChar
                        {
                            Font = "Wingdings",
                            Char = "F0E0"
                        })));
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        result.ApplyTo(model);

        const string tab = "<span style=\"white-space:pre\">&#9;</span>";
        await Assert.That(model.Body)
            .IsEqualTo($"<p>a{tab}b<strong>{tab}</strong>co‑ops<br>•</p>");
    }

    [Test]
    public async Task HtmlEditableRendersPlaceholderWhenEmpty()
    {
        var model = new EditableArticle
        {
            Title = "T",
            Body = ""
        };
        using var stream = await RenderModel("{{ Body }}", model);

        using var doc = WordprocessingDocument.Open(stream, false);
        var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
        await Assert.That(sdt.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();
        await Assert.That(sdt.InnerText).IsEqualTo("Enter rich text");

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableArticle>(stream);
        var field = result.Fields.First(_ => _.Path == "Body");
        await Assert.That(field.State).IsEqualTo(FieldState.Empty);
    }


    [ParchmentBindable]
    public partial class EditableQuote
    {
        [EditableField]
        public required string Reference { get; set; }

        public required EditableCustomer Customer;
    }

    public class EditableCustomer
    {
        [EditableField]
        public required string ContactEmail { get; set; }
    }

    [ParchmentBindable]
    public partial class EditableAccount
    {
        public EditableProfile? Profile { get; set; }
    }

    public class EditableProfile
    {
        [EditableField]
        public bool Active { get; set; }

        [EditableField]
        public int Seats { get; set; }

        [EditableField]
        public Date Joined { get; set; }

        [EditableField]
        public QuoteStatus Status { get; set; }

        [EditableField]
        public string Name { get; set; } = "";
    }

    const string nullParentTemplate =
        """
        Active: {{ Profile.Active }}

        Seats: {{ Profile.Seats }}

        Joined: {{ Profile.Joined }}

        Status: {{ Profile.Status }}

        Name: {{ Profile.Name }}
        """;

    static async Task<MemoryStream> Render(string templateContent, EditableOrder model)
    {
        using var template = DocxTemplateBuilder.Build(templateContent);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static async Task<MemoryStream> RenderModel<T>(string templateContent, T model)
    {
        using var template = DocxTemplateBuilder.Build(templateContent);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);

        var stream = new MemoryStream();
        await store.Render(model!, stream);
        stream.Position = 0;
        return stream;
    }

    static SdtRun FindSdt(OpenXmlCompositeElement root, string tag)
    {
        var sdt = root.Descendants<SdtRun>()
            .FirstOrDefault(_ => _.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag);
        if (sdt == null)
        {
            throw new InvalidOperationException($"No sdt with tag '{tag}' found");
        }

        return sdt;
    }

    static SdtBlock FindSdtBlock(OpenXmlCompositeElement root, string tag)
    {
        var sdt = root.Descendants<SdtBlock>()
            .FirstOrDefault(_ => _.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag);
        if (sdt == null)
        {
            throw new InvalidOperationException($"No block sdt with tag '{tag}' found");
        }

        return sdt;
    }

    static MemoryStream BuildStraddledTemplate() =>
        BuildParagraphTemplate(
            new Run(Preserved("PO: {{ Purch")),
            new Run(
                new RunProperties(new Bold()),
                Preserved("aseOrder }}")),
            new Run(Preserved(" end")));

    static Text Preserved(string text) =>
        new(text)
        {
            Space = SpaceProcessingModeValues.Preserve
        };

    // Renders the single-paragraph template and flattens the paragraph into its content order:
    // text as-is, each content control as [value], tabs and breaks as <tab> / <br>.
    static async Task<string> RenderLayout(MemoryStream template)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();
        var builder = new StringBuilder();
        foreach (var element in paragraph.Descendants())
        {
            switch (element)
            {
                case SdtRun sdt:
                    builder.Append($"[{sdt.InnerText}]");
                    break;
                case Text text when !text.Ancestors<SdtRun>().Any():
                    builder.Append(text.Text);
                    break;
                case TabChar:
                    builder.Append("<tab>");
                    break;
                case Break:
                    builder.Append("<br>");
                    break;
            }
        }

        return builder.ToString();
    }

    // As RenderLayout, but nested: an inline wrapper shows as <name>...</name> around what it
    // holds, and a control's editable range has to open and close in the control's own
    // container. Also checks the output is schema-valid and the control still extracts.
    static async Task<string> RenderWrapperLayout(MemoryStream template)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(template);

        using var stream = new MemoryStream();
        await store.Render(NewOrder(), stream);
        stream.Position = 0;

        var builder = new StringBuilder();
        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
                .Validate(doc)
                .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
                .ToList();
            await Assert.That(errors).IsEmpty();

            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();
            foreach (var sdt in paragraph.Descendants<SdtRun>().Where(_ => _.SdtProperties?.GetFirstChild<SdtLock>() != null))
            {
                await Assert.That(sdt.PreviousSibling()).IsTypeOf<PermStart>();
                await Assert.That(sdt.NextSibling()).IsTypeOf<PermEnd>();
            }

            AppendWrapperLayout(paragraph, builder);
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<EditableOrder>(stream);
        var field = result.Fields.Single(_ => _.Path == "PurchaseOrder");
        await Assert.That(field.Value).IsEqualTo("PO-2026-17");

        return builder.ToString();
    }

    static void AppendWrapperLayout(OpenXmlElement container, StringBuilder builder)
    {
        foreach (var element in container.ChildElements)
        {
            switch (element)
            {
                case SdtRun sdt when sdt.SdtProperties?.GetFirstChild<SdtLock>() != null:
                    builder.Append($"[{sdt.InnerText}]");
                    break;
                case SdtRun sdt:
                    builder.Append("<sdt>");
                    AppendWrapperLayout(sdt.SdtContentRun!, builder);
                    builder.Append("</sdt>");
                    break;
                case Run run:
                    builder.Append(run.InnerText);
                    break;
                case OpenXmlCompositeElement composite when composite is not ParagraphProperties:
                    builder.Append($"<{composite.LocalName}>");
                    AppendWrapperLayout(composite, builder);
                    builder.Append($"</{composite.LocalName}>");
                    break;
            }
        }
    }

    static MemoryStream BuildParagraphTemplate(params OpenXmlElement[] children) =>
        BuildBodyTemplate(new Paragraph(children));

    static MemoryStream BuildBodyTemplate(params OpenXmlElement[] blocks)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body(blocks);
            body.AppendChild(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));
            mainPart.Document = new(body);
        }

        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildTemplateWithHeader(string bodyText, string headerText)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(
                new Paragraph(
                    new Run(
                        new Text(headerText)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        })));

            body.Append(
                new Paragraph(
                    new Run(
                        new Text(bodyText)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        })));
            body.Append(
                new SectionProperties(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    },
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));
        }

        stream.Position = 0;
        return stream;
    }
}
