using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
// ReSharper disable PartialTypeWithSinglePart

public partial class EditableCollectionTests
{
    #region EditableCollectionModel
    public class Budget
    {
        [EditableField]
        public required string Year { get; set; }

        [EditableField]
        public required decimal Amount { get; set; }
    }

    [ParchmentBindable]
    public partial class BudgetPlan
    {
        public required string Title;

        [EditableField]
        public required List<Budget> Budgets { get; set; }
    }
    #endregion

    public enum Priority
    {
        Low,
        High
    }

    public class Chore
    {
        [EditableField]
        public string? Name { get; set; }

        [EditableField]
        public bool Done { get; set; }

        [EditableField]
        public int Count { get; set; }

        [EditableField]
        public Date Due { get; set; }

        [EditableField]
        public Priority Priority { get; set; }

        [EditableField]
        public DateTimeOffset Logged { get; set; }

        [EditableField]
        public Time At { get; set; }
    }

    [ParchmentBindable]
    public partial class ChoreList
    {
        [EditableField]
        public required List<Chore> Chores { get; set; }
    }

    const string choreTemplate =
        """
        {% for c in Chores %}

        {{ c.Name }} {{ c.Done }} {{ c.Count }} {{ c.Due }} {{ c.Priority }} {{ c.Logged }} {{ c.At }}

        {% endfor %}
        """;

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
        using var stream = await Render(template, NewPlan());

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
        using var stream = await Render(template, model);

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
    public async Task BlankItemInCharacterStyledRunValidates()
    {
        using var docx = DocxTemplateBuilder.Build(template);
        using (var doc = WordprocessingDocument.Open(docx, true))
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

        docx.Position = 0;
        var store = new TemplateStore();
        store.RegisterDocxTemplate<BudgetPlan>(docx);

        using var stream = new MemoryStream();
        await store.Render(
            new BudgetPlan
            {
                Title = "Empty",
                Budgets = []
            },
            stream);
        stream.Position = 0;

        using var rendered = WordprocessingDocument.Open(stream, false);
        var container = RepeatingSection(rendered.MainDocumentPart!.Document!.Body!, "Budgets");

        // The blank item's placeholder run swaps the site's character style for the placeholder one.
        var year = container.Descendants<SdtRun>().First(_ => Tag(_) == "Year");
        await Assert.That(year.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();
        var styles = year.Descendants<Run>().Single().RunProperties!.Elements<RunStyle>().ToList();
        await Assert.That(styles.Count).IsEqualTo(1);
        await Assert.That(styles[0].Val!.Value).IsEqualTo("PlaceholderText");

        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(rendered)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task OutputValidates()
    {
        using var stream = await Render(template, NewPlan());

        using var doc = WordprocessingDocument.Open(stream, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator.Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }


    [Test]
    public async Task RoundTripsCollection()
    {
        using var stream = await Render(template, NewPlan());

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
        using var stream = await Render(template, NewPlan());

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

    [Test]
    public async Task EmptyCollectionRoundTripsAsEmpty()
    {
        using var stream = await Render(
            template,
            new BudgetPlan
            {
                Title = "Empty",
                Budgets = []
            });

        // The blank row carries no default values: a non-nullable member shows placeholder text
        // rather than "0", so the untouched row extracts as nothing.
        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var amount = Control(doc.MainDocumentPart!.Document!.Body!, "Amount");
            await Assert.That(amount.SdtProperties!.GetFirstChild<ShowingPlaceholder>()).IsNotNull();
        }

        stream.Position = 0;
        var model = NewPlan();
        ParchmentExtractor.Extract<BudgetPlan>(stream).ApplyTo(model);

        await Assert.That(model.Budgets.Count).IsEqualTo(0);
    }

    [Test]
    public async Task EmptyCollectionOfEveryKindRoundTripsAsEmpty()
    {
        using var stream = await Render(
            choreTemplate,
            new ChoreList
            {
                Chores = []
            });

        using (var doc = WordprocessingDocument.Open(stream, false))
        {
            var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013);
            var errors = validator.Validate(doc)
                .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
                .ToList();
            await Assert.That(errors).IsEmpty();
        }

        stream.Position = 0;
        var model = new ChoreList
        {
            Chores = [new()]
        };
        ParchmentExtractor.Extract<ChoreList>(stream).ApplyTo(model);

        await Assert.That(model.Chores.Count).IsEqualTo(0);
    }

    [Test]
    public async Task FilledBlankRowExtractsOneItem()
    {
        using var stream = await Render(
            template,
            new BudgetPlan
            {
                Title = "Empty",
                Budgets = []
            });

        // Only Year is filled in; Amount stays blank and keeps the element's default.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            SetText(Control(doc.MainDocumentPart!.Document!.Body!, "Year"), "2030");
            doc.Save();
        }

        stream.Position = 0;
        var model = NewPlan();
        ParchmentExtractor.Extract<BudgetPlan>(stream).ApplyTo(model);

        await Assert.That(model.Budgets.Count).IsEqualTo(1);
        await Assert.That(model.Budgets[0].Year).IsEqualTo("2030");
        await Assert.That(model.Budgets[0].Amount).IsEqualTo(0m);
    }

    [Test]
    public async Task CheckedBlankRowExtractsOneItem()
    {
        using var stream = await Render(
            choreTemplate,
            new ChoreList
            {
                Chores = []
            });

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var done = Control(doc.MainDocumentPart!.Document!.Body!, "Done");
            done.SdtProperties!
                .GetFirstChild<W14.SdtContentCheckBox>()!
                .GetFirstChild<W14.Checked>()!
                .Val = W14.OnOffValues.One;
            doc.Save();
        }

        stream.Position = 0;
        var model = new ChoreList
        {
            Chores = []
        };
        ParchmentExtractor.Extract<ChoreList>(stream).ApplyTo(model);

        await Assert.That(model.Chores.Count).IsEqualTo(1);
        await Assert.That(model.Chores[0].Done).IsTrue();
        await Assert.That(model.Chores[0].Count).IsEqualTo(0);
    }

    [Test]
    public async Task AllDefaultItemRoundTrips()
    {
        // A row rendered from a real item is kept even when every value is its default: only the
        // blank row of an empty collection is dropped for being untouched.
        using var budgets = await Render(
            template,
            new BudgetPlan
            {
                Title = "Zero",
                Budgets =
                [
                    new()
                    {
                        Year = "",
                        Amount = 0m
                    }
                ]
            });
        var plan = NewPlan();
        ParchmentExtractor.Extract<BudgetPlan>(budgets).ApplyTo(plan);
        await Assert.That(plan.Budgets.Count).IsEqualTo(1);
        await Assert.That(plan.Budgets[0].Amount).IsEqualTo(0m);

        using var chores = await Render(
            choreTemplate,
            new ChoreList
            {
                Chores = [new()]
            });
        var list = new ChoreList
        {
            Chores = []
        };
        ParchmentExtractor.Extract<ChoreList>(chores).ApplyTo(list);
        await Assert.That(list.Chores.Count).IsEqualTo(1);
        await Assert.That(list.Chores[0].Done).IsFalse();
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

    static async Task<MemoryStream> Render<T>(string templateContent, T model)
    {
        using var docx = DocxTemplateBuilder.Build(templateContent);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(docx);

        var stream = new MemoryStream();
        await store.Render(model!, stream);
        stream.Position = 0;
        return stream;
    }
}
