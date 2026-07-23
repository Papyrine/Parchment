using W15 = DocumentFormat.OpenXml.Office2013.Word;

public class EditableCollectionTests
{
    #region EditableCollectionModel
    public class Budget
    {
        [EditableField]
        public required string Year { get; set; }

        [EditableField]
        public required decimal Amount { get; set; }
    }

    public class BudgetPlan
    {
        public required string Title;

        [EditableField]
        public required List<Budget> Budgets { get; set; }
    }
    #endregion

    static BudgetPlan NewPlan() =>
        new()
        {
            Title = "Plan A",
            Budgets =
            [
                new()
                {
                    Year = "2026",
                    Amount = 10m
                },
                new()
                {
                    Year = "2027",
                    Amount = 20m
                }
            ]
        };

    const string template =
        """
        {% for b in Budgets %}

        Year: {{ b.Year }} Amount: {{ b.Amount }}

        {% endfor %}
        """;

    [Test]
    public async Task RendersRepeatingSection()
    {
        using var stream = await Render(template, NewPlan(), "budgets");

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var container = RepeatingSection(body, "Budgets");
        var itemsSdt = container.Descendants<SdtBlock>()
            .Where(_ => _.SdtProperties?.GetFirstChild<W15.SdtRepeatedSectionItem>() != null)
            .ToList();
        await Assert.That(itemsSdt.Count).IsEqualTo(2);

        // Each item carries item-relative controls (tags Year / Amount) with the item's values.
        var firstYear = itemsSdt[0].Descendants<SdtRun>().First(_ => Tag(_) == "Year");
        await Assert.That(firstYear.InnerText).IsEqualTo("2026");
        var firstAmount = itemsSdt[0].Descendants<SdtRun>().First(_ => Tag(_) == "Amount");
        await Assert.That(firstAmount.InnerText).IsEqualTo("10");
        var secondYear = itemsSdt[1].Descendants<SdtRun>().First(_ => Tag(_) == "Year");
        await Assert.That(secondYear.InnerText).IsEqualTo("2027");

        // Wrapped in a single editable perm range around the container.
        await Assert.That(body.Descendants<PermStart>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<PermEnd>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task EmptyCollectionRendersOneBlankItem()
    {
        var model = new BudgetPlan
        {
            Title = "Empty",
            Budgets = []
        };
        using var stream = await Render(template, model, "budgets-empty");

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var container = RepeatingSection(body, "Budgets");
        var itemsSdt = container.Descendants<SdtBlock>()
            .Where(_ => _.SdtProperties?.GetFirstChild<W15.SdtRepeatedSectionItem>() != null)
            .ToList();
        // One blank item so Word has a clone template.
        await Assert.That(itemsSdt.Count).IsEqualTo(1);
        var year = itemsSdt[0].Descendants<SdtRun>().First(_ => Tag(_) == "Year");
        await Assert.That(year.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();
    }

    [Test]
    public async Task OutputValidates()
    {
        using var stream = await Render(template, NewPlan(), "budgets-valid");

        using var doc = WordprocessingDocument.Open(stream, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    public class EmptyElement
    {
        public string Name { get; set; } = "";
    }

    public class NoEditableElementModel
    {
        [EditableField]
        public required List<EmptyElement> Items { get; set; }
    }

    [Test]
    public async Task ElementWithoutEditableMembersIsRejected()
    {
        using var docx = DocxTemplateBuilder.Build("x");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<NoEditableElementModel>("no-editable-element", docx))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("no [EditableField] members");
    }

    [Test]
    public async Task RoundTripsCollection()
    {
        using var stream = await Render(template, NewPlan(), "budgets-rt");

        var model = new BudgetPlan
        {
            Title = "x",
            Budgets = []
        };

        #region EditableCollectionExtract
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        // model.Budgets is the edited list - added rows appear, removed rows are gone
        result.ApplyTo(model);
        #endregion

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(model.Budgets.Count).IsEqualTo(2);
        await Assert.That(model.Budgets[0].Year).IsEqualTo("2026");
        await Assert.That(model.Budgets[0].Amount).IsEqualTo(10m);
        await Assert.That(model.Budgets[1].Year).IsEqualTo("2027");
        await Assert.That(model.Budgets[1].Amount).IsEqualTo(20m);
    }

    [Test]
    public async Task RoundTripsAfterEditAddRemove()
    {
        using var stream = await Render(template, NewPlan(), "budgets-edit");

        // Simulate Word: edit row 0, delete row 1, add a cloned row.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var content = RepeatingSection(body, "Budgets").ChildElements.First(_ => _.LocalName == "sdtContent");
            var items = content.ChildElements.Where(IsItem).Cast<SdtBlock>().ToList();

            SetText(Control(items[0], "Year"), "2099");
            items[1].Remove();

            var clone = (SdtBlock)items[0].CloneNode(true);
            SetText(Control(clone, "Year"), "2100");
            SetText(Control(clone, "Amount"), "30");
            content.AppendChild(clone);
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);
        var model = new BudgetPlan
        {
            Title = "x",
            Budgets = []
        };
        result.ApplyTo(model);

        await Assert.That(model.Budgets.Count).IsEqualTo(2);
        await Assert.That(model.Budgets[0].Year).IsEqualTo("2099");
        await Assert.That(model.Budgets[0].Amount).IsEqualTo(10m);
        await Assert.That(model.Budgets[1].Year).IsEqualTo("2100");
        await Assert.That(model.Budgets[1].Amount).IsEqualTo(30m);
    }

    static bool IsItem(OpenXmlElement element) =>
        element is SdtBlock sdt &&
        sdt.SdtProperties?.GetFirstChild<W15.SdtRepeatedSectionItem>() != null;

    static SdtElement Control(OpenXmlElement item, string tag) =>
        item.Descendants<SdtElement>().First(_ => Tag(_) == tag);

    static void SetText(SdtElement sdt, string text)
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

    static SdtBlock RepeatingSection(OpenXmlCompositeElement root, string tag) =>
        root.Descendants<SdtBlock>()
            .First(_ => _.SdtProperties?.GetFirstChild<W15.SdtRepeatedSection>() != null &&
                        Tag(_) == tag);

    static string? Tag(SdtElement sdt) =>
        sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    static async Task<MemoryStream> Render<T>(string templateContent, T model, string name)
    {
        using var docx = DocxTemplateBuilder.Build(templateContent);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(name, docx);

        var stream = new MemoryStream();
        await store.Render(name, model!, stream);
        stream.Position = 0;
        return stream;
    }
}
