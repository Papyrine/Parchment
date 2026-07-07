using DocumentFormat.OpenXml.Validation;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using SdtLock = DocumentFormat.OpenXml.Wordprocessing.Lock;

public class EditableFieldTests
{
    static string ScenarioPath(string scenarioName) =>
        Path.Combine(
            ProjectFiles.ProjectDirectory,
            "Scenarios",
            scenarioName);

    #region EditableFieldsModel
    public class OrderForm
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
        store.RegisterDocxTemplate<OrderForm>("order-form", templatePath);

        var model = new OrderForm
        {
            Number = "ORD-2026-042",
            PurchaseOrder = "PO-77041",
            Approved = true,
            Status = OrderStatus.Submitted,
            Notes = null
        };

        using var stream = new MemoryStream();
        await store.Render("order-form", model, stream);
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
        store.RegisterDocxTemplate<OrderForm>("order-form-roundtrip", templatePath);

        var model = new OrderForm
        {
            Number = "ORD-2026-042",
            PurchaseOrder = "PO-77041",
            Approved = false,
            Status = OrderStatus.Draft,
            Notes = null
        };

        using var stream = new MemoryStream();
        await store.Render("order-form-roundtrip", model, stream);
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

    public class EditableOrder
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
        // No {{ Delivery }} here: Morph's page render formats the date control from w:fullDate
        // with the machine locale, which would make the png snapshot culture-dependent. Date
        // structure is asserted in SdtStructurePerKind instead.
        using var template = DocxTemplateBuilder.Build(
            """
            Order {{ Number }}

            PO: {{ PurchaseOrder }}

            Approved: {{ Approved }}

            Status: {{ Status }}

            Discount: {{ Discount }}

            Notes: {{ Notes }}

            {{ Instructions }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>("editable-all", template);

        using var stream = new MemoryStream();
        await store.Render("editable-all", NewOrder(), stream);

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
        await Assert.That(date!.FullDate!.Value).IsEqualTo(new DateTime(2026, 7, 6));
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
        await Assert.That(date!.FullDate!.Value).IsEqualTo(new DateTime(2026, 7, 6, 14, 30, 0));
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
        store.RegisterDocxTemplate<EditableOrder>("editable-unprotected", template, ProtectionMode.None);

        using var stream = new MemoryStream();
        await store.Render("editable-unprotected", NewOrder(), stream);
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
        store.RegisterDocxTemplate<Invoice>("no-editable", template);

        using var stream = new MemoryStream();
        await store.Render("no-editable", SampleData.Invoice(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var protection = doc.MainDocumentPart!.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentProtection>();
        await Assert.That(protection).IsNull();
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
        store.RegisterDocxTemplate<EditableOrder>("editable-straddle", template);

        using var stream = new MemoryStream();
        await store.Render("editable-straddle", NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).IsEqualTo("PO: PO-2026-17 end");
        await Assert.That(FindSdt(body, "PurchaseOrder").InnerText).IsEqualTo("PO-2026-17");
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
        store.RegisterDocxTemplate<EditableQuote>("editable-nested", template);

        var model = new EditableQuote
        {
            Reference = "Q-7",
            Customer = new()
            {
                ContactEmail = "ada@example.com"
            }
        };
        using var stream = new MemoryStream();
        await store.Render("editable-nested", model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(FindSdt(body, "Customer.ContactEmail").InnerText).IsEqualTo("ada@example.com");
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
        store.RegisterDocxTemplate<EditableOrder>("editable-header", template);

        using var stream = new MemoryStream();
        await store.Render("editable-header", NewOrder(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Count()).IsEqualTo(1);

        var header = doc.MainDocumentPart.HeaderParts.Single().Header!;
        await Assert.That(header.Descendants<SdtRun>().Any()).IsFalse();
        await Assert.That(header.InnerText).IsEqualTo("PO: PO-2026-17");
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

    [Test]
    public async Task FilterOnEditableTokenIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("{{ PurchaseOrder | upcase }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EditableOrder>("editable-filter", template))
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
                () => store.RegisterDocxTemplate<EditableOrder>("editable-duplicate", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }

    public class UnsupportedTypeModel
    {
        [EditableField]
        public List<string> Items { get; set; } = [];
    }

    [Test]
    public async Task UnsupportedMemberTypeIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("x");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<UnsupportedTypeModel>("editable-unsupported", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("unsupported type");
    }

    public class NullableBoolModel
    {
        [EditableField]
        public bool? Maybe { get; set; }
    }

    [Test]
    public async Task NullableBoolIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("x");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<NullableBoolModel>("editable-nullable-bool", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("bool?");
    }

    public class InitOnlyModel
    {
        [EditableField]
        public string PurchaseOrder { get; init; } = "";
    }

    [Test]
    public async Task InitOnlySetterIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("x");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<InitOnlyModel>("editable-init-only", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("setter");
    }

    public class GetOnlyModel
    {
        public string Prefix = "x";

        [EditableField]
        public string Computed => Prefix;
    }

    [Test]
    public async Task GetOnlyMemberIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("x");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<GetOnlyModel>("editable-get-only", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("setter");
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    public class ConflictModel
    {
        [EditableField]
        [Markdown]
        public string Body { get; set; } = "";
    }

    [Test]
    public async Task ConflictingFormatAttributeIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("x");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<ConflictModel>("editable-conflict", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("cannot be combined");
    }

    public class MixedStructuralModel
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
                () => store.RegisterDocxTemplate<MixedStructuralModel>("editable-mixed", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("own paragraph");
    }

    public class EditableQuote
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

    static async Task<MemoryStream> Render(string templateContent, object model, [CallerMemberName] string name = "")
    {
        using var template = DocxTemplateBuilder.Build(templateContent);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableOrder>(name, template);

        var stream = new MemoryStream();
        await store.Render(name, model, stream);
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

    static MemoryStream BuildStraddledTemplate()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(
                new Body(
                    new Paragraph(
                        new Run(
                            new Text("PO: {{ Purch")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }),
                        new Run(
                            new RunProperties(new Bold()),
                            new Text("aseOrder }}")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }),
                        new Run(
                            new Text(" end")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            })),
                    new SectionProperties(
                        new PageSize
                        {
                            Width = 6500,
                            Height = 8000
                        })));
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
