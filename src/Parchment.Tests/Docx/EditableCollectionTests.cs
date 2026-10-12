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

    public class ChecklistItem
    {
        [EditableField]
        public string Label { get; set; } = "";

        [EditableField]
        public bool Done { get; set; }
    }

    [ParchmentBindable]
    public partial class Checklist
    {
        [EditableField]
        public List<ChecklistItem> Items { get; set; } = [];
    }

    [Test]
    public async Task NullItemRendersUncheckedCheckbox()
    {
        var model = new Checklist
        {
            Items =
            [
                new()
                {
                    Label = "First",
                    Done = true
                },
                null!
            ]
        };
        using var stream = await Render(
            """
            {% for i in Items %}

            {{ i.Label }}: {{ i.Done }}

            {% endfor %}
            """,
            model);

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        var items = RepeatingSection(body, "Items").Descendants<SdtBlock>().Where(IsItem).ToList();
        await Assert.That(items.Count).IsEqualTo(2);
        await Assert.That(Control(items[0], "Done").InnerText).IsEqualTo("☒");

        // The null item's getters all yield null: placeholder for the text, an unchecked box.
        await Assert.That(Control(items[1], "Label").InnerText).IsEqualTo("Enter text");
        await Assert.That(Control(items[1], "Done").InnerText).IsEqualTo("☐");

        // A null item in a populated list is not the blank row of an empty collection: it carries
        // no blank-item tag, so its unchecked box is a value and the row survives extraction.
        await Assert.That(Tag(items[0])).IsNull();
        await Assert.That(Tag(items[1])).IsNull();

        stream.Position = 0;
        var extracted = new Checklist();
        ParchmentExtractor.Extract<Checklist>(stream).ApplyTo(extracted);
        await Assert.That(extracted.Items.Count).IsEqualTo(2);
        await Assert.That(extracted.Items[0].Label).IsEqualTo("First");
        await Assert.That(extracted.Items[0].Done).IsTrue();
        await Assert.That(extracted.Items[1].Done).IsFalse();
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

    public class Depot
    {
        [EditableField]
        public string? Code { get; set; }
    }

    public class Shipping
    {
        [EditableField]
        public string? Note { get; set; }

        [EditableField]
        public bool Fragile { get; set; }

        [EditableField]
        public int Boxes { get; set; }

        [EditableField]
        public Date Due { get; set; }

        [EditableField]
        public Priority Priority { get; set; }

        [EditableField]
        public DateTimeOffset Logged { get; set; }

        [EditableField]
        public Time At { get; set; }

        [EditableField]
        [System.Diagnostics.CodeAnalysis.StringSyntax("html")]
        public string? Instructions { get; set; }

        public Depot? Depot { get; set; }
    }

    public class Parcel
    {
        [EditableField]
        public string? Name { get; set; }

        public Shipping? Shipping { get; set; }
    }

    [ParchmentBindable]
    public partial class Manifest
    {
        [EditableField]
        public List<Parcel> Parcels { get; set; } = [];
    }

    const string manifestTemplate =
        """
        {% for p in Parcels %}

        {{ p.Name }} {{ p.Shipping.Note }} {{ p.Shipping.Fragile }} {{ p.Shipping.Boxes }} {{ p.Shipping.Due }} {{ p.Shipping.Priority }} {{ p.Shipping.Logged }} {{ p.Shipping.At }} {{ p.Shipping.Depot.Code }}

        {{ p.Shipping.Instructions }}

        {% endfor %}
        """;

    static Manifest NewManifest() =>
        new()
        {
            Parcels =
            [
                new()
                {
                    Name = "first",
                    Shipping = new()
                    {
                        Note = "leave at door",
                        Fragile = true,
                        Boxes = 3,
                        Due = new(2026, 3, 4),
                        Priority = Priority.High,
                        Logged = new(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(2)),
                        At = new(9, 30),
                        Instructions = "<p>Keep <strong>upright</strong></p>",
                        Depot = new()
                        {
                            Code = "SYD"
                        }
                    }
                },
                new()
                {
                    Name = "second"
                }
            ]
        };

    // Extraction builds each item from nothing, so the objects between the item and a nested
    // editable member have to be created on the way to the value, at every depth and for every
    // kind of control.
    [Test]
    public async Task NestedElementMembersRoundTrip()
    {
        using var stream = await Render(manifestTemplate, NewManifest());

        var result = ParchmentExtractor.Extract<Manifest>(stream);
        var model = new Manifest();
        result.ApplyTo(model);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(model.Parcels.Count).IsEqualTo(2);

        var shipping = model.Parcels[0].Shipping!;
        await Assert.That(shipping.Note).IsEqualTo("leave at door");
        await Assert.That(shipping.Fragile).IsTrue();
        await Assert.That(shipping.Boxes).IsEqualTo(3);
        await Assert.That(shipping.Due).IsEqualTo(new Date(2026, 3, 4));
        await Assert.That(shipping.Priority).IsEqualTo(Priority.High);
        await Assert.That(shipping.Logged).IsEqualTo(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(2)));
        await Assert.That(shipping.At).IsEqualTo(new Time(9, 30));
        await Assert.That(shipping.Instructions).IsEqualTo("<p>Keep <strong>upright</strong></p>");
        await Assert.That(shipping.Depot!.Code).IsEqualTo("SYD");

        await Assert.That(model.Parcels[1].Name).IsEqualTo("second");
    }

    [Test]
    public async Task NestedElementMemberEditedInWordIsExtracted()
    {
        using var stream = await Render(manifestTemplate, NewManifest());

        // The second row was rendered from an item with no Shipping at all: its nested controls
        // show placeholders, and the user fills two of them in.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var content = RepeatingSection(body, "Parcels").ChildElements.First(_ => _.LocalName == "sdtContent");
            var items = content.ChildElements.Where(IsItem).Cast<SdtBlock>().ToList();

            SetText(Control(items[0], "Shipping.Note"), "ring the bell");
            SetText(Control(items[1], "Shipping.Note"), "new note");
            SetText(Control(items[1], "Shipping.Depot.Code"), "MEL");
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<Manifest>(stream);
        var model = new Manifest();
        result.ApplyTo(model);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(model.Parcels[0].Shipping!.Note).IsEqualTo("ring the bell");
        await Assert.That(model.Parcels[1].Shipping!.Note).IsEqualTo("new note");
        await Assert.That(model.Parcels[1].Shipping!.Depot!.Code).IsEqualTo("MEL");
    }

    public abstract class Carrier
    {
        [EditableField]
        public string? Contact { get; set; }
    }

    public class Courier : Carrier;

    public class Seal(string code)
    {
        public string Code { get; } = code;

        [EditableField]
        public string? Colour { get; set; }
    }

    public class Customs
    {
        [EditableField]
        public string? Tariff { get; set; }
    }

    public class Crate
    {
        [EditableField]
        public string? Name { get; set; }

        public Carrier? Carrier { get; set; }

        public Seal? Seal { get; set; }

        public Customs Customs { get; } = new();
    }

    [ParchmentBindable]
    public partial class Hold
    {
        [EditableField]
        public List<Crate> Crates { get; set; } = [];
    }

    // An abstract type, or one with no parameterless constructor, cannot be created for a new item,
    // so the value has nowhere to go. That is reported instead of being dropped; a parent the
    // element initializes itself needs no creating and still round-trips.
    [Test]
    public async Task NestedElementMemberThatCannotBeCreatedIsReported()
    {
        var crate = new Crate
        {
            Name = "one",
            Carrier = new Courier
            {
                Contact = "Ann"
            },
            Seal = new("s1")
            {
                Colour = "red"
            }
        };
        crate.Customs.Tariff = "8471";
        using var stream = await Render(
            """
            {% for c in Crates %}

            {{ c.Name }} {{ c.Carrier.Contact }} {{ c.Seal.Colour }} {{ c.Customs.Tariff }}

            {% endfor %}
            """,
            new Hold
            {
                Crates = [crate, crate]
            });

        var result = ParchmentExtractor.Extract<Hold>(stream);
        var model = new Hold();
        result.ApplyTo(model);

        await Assert.That(result.AllExtracted).IsFalse();
        var unreachable = result.Fields
            .Where(_ => _.State == FieldState.Unreachable)
            .Select(_ => $"{_.Path}={_.Value}")
            .ToList();
        await Assert.That(unreachable).IsEquivalentTo(["Crates.Carrier.Contact=Ann", "Crates.Seal.Colour=red"]);

        await Assert.That(model.Crates.Count).IsEqualTo(2);
        await Assert.That(model.Crates[0].Name).IsEqualTo("one");
        await Assert.That(model.Crates[0].Customs.Tariff).IsEqualTo("8471");
        await Assert.That(model.Crates[0].Carrier).IsNull();
        await Assert.That(model.Crates[0].Seal).IsNull();
    }

    // Extraction is replace-all: the rows in the document become the collection. A loop that
    // rendered only a window of it, or rendered it backwards, would have that written back as the
    // data, so the modifiers are refused here instead of applied.
    [Test]
    [Arguments("limit:1")]
    [Arguments("offset:1")]
    [Arguments("reversed")]
    public async Task LoopModifierFailsRegistration(string modifier)
    {
        using var docx = DocxTemplateBuilder.Build(
            $$$"""
               {% for b in Budgets {{{modifier}}} %}

               Year: {{ b.Year }} Amount: {{ b.Amount }}

               {% endfor %}
               """);
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<BudgetPlan>(docx))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("Budgets");
        await Assert.That(exception.Message).Contains("[EditableField]");
    }

    // A loop nested in another renders read-only — nothing it shows is extracted — so a modifier
    // there is applied like on any other loop.
    [Test]
    public async Task LoopModifierOnNestedReadOnlyLoopIsApplied()
    {
        using var stream = await Render(
            """
            {% for b in Budgets %}

            Year: {{ b.Year }}

            {% for other in Budgets limit:1 %}

            first={{ other.Year }}

            {% endfor %}

            {% endfor %}
            """,
            NewPlan());

        using var doc = WordprocessingDocument.Open(stream, false);
        var nested = doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.StartsWith("first=", StringComparison.Ordinal))
            .ToList();
        await Assert.That(nested).IsEquivalentTo(["first=2026", "first=2026"]);
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
