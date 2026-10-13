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

    [ParchmentBindable]
    public partial class BudgetLedger
    {
        [EditableField]
        public required Budget[] Budgets { get; set; }
    }

    // One result applied to two models: each gets its own collection and its own items, so
    // changing one model leaves the other, and a later ApplyTo, as the document had it.
    [Test]
    public async Task ApplyToGivesEachModelItsOwnList()
    {
        using var stream = await Render(template, NewPlan());
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        var first = new BudgetPlan
        {
            Title = "first",
            Budgets = []
        };
        var second = new BudgetPlan
        {
            Title = "second",
            Budgets = []
        };
        result.ApplyTo(first);
        result.ApplyTo(second);

        await Assert.That(ReferenceEquals(first.Budgets, second.Budgets)).IsFalse();
        await Assert.That(ReferenceEquals(first.Budgets[0], second.Budgets[0])).IsFalse();

        first.Budgets[0].Year = "1999";
        first.Budgets.RemoveAt(1);
        first.Budgets.Add(
            new()
            {
                Year = "2100",
                Amount = 30m
            });

        await Assert.That(second.Budgets.Count).IsEqualTo(2);
        await Assert.That(second.Budgets[0].Year).IsEqualTo("2026");
        await Assert.That(second.Budgets[1].Year).IsEqualTo("2027");

        var third = new BudgetPlan
        {
            Title = "third",
            Budgets = []
        };
        result.ApplyTo(third);
        await Assert.That(third.Budgets.Count).IsEqualTo(2);
        await Assert.That(third.Budgets[0].Year).IsEqualTo("2026");
        await Assert.That(third.Budgets[0].Amount).IsEqualTo(10m);
        await Assert.That(third.Budgets[1].Year).IsEqualTo("2027");
        await Assert.That(third.Budgets[1].Amount).IsEqualTo(20m);
    }

    [Test]
    public async Task ApplyToGivesEachModelItsOwnArray()
    {
        using var stream = await Render(
            template,
            new BudgetLedger
            {
                Budgets = NewPlan().Budgets.ToArray()
            });
        var result = ParchmentExtractor.Extract<BudgetLedger>(stream);

        var first = new BudgetLedger
        {
            Budgets = []
        };
        var second = new BudgetLedger
        {
            Budgets = []
        };
        result.ApplyTo(first);
        result.ApplyTo(second);

        await Assert.That(ReferenceEquals(first.Budgets, second.Budgets)).IsFalse();
        await Assert.That(ReferenceEquals(first.Budgets[0], second.Budgets[0])).IsFalse();

        first.Budgets[0].Year = "1999";
        first.Budgets[1] = first.Budgets[0];

        await Assert.That(second.Budgets.Length).IsEqualTo(2);
        await Assert.That(second.Budgets[0].Year).IsEqualTo("2026");
        await Assert.That(second.Budgets[0].Amount).IsEqualTo(10m);
        await Assert.That(second.Budgets[1].Year).IsEqualTo("2027");
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

    [Test]
    public async Task RowFieldsAreReported()
    {
        using var stream = await Render(template, NewPlan());

        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(result.Fields.Select(_ => _.Path))
            .IsEquivalentTo(["Budgets[0].Year", "Budgets[0].Amount", "Budgets[1].Year", "Budgets[1].Amount"]);
        var amount = result.Fields.Single(_ => _.Path == "Budgets[1].Amount");
        await Assert.That(amount.State).IsEqualTo(FieldState.Extracted);
        await Assert.That(amount.Value).IsEqualTo(20m);
    }

    [Test]
    public async Task UnparseableRowValueIsReported()
    {
        using var stream = await Render(template, NewPlan());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var items = Items(doc, "Budgets");
            SetText(Control(items[1], "Amount"), "abc");
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        await Assert.That(result.AllExtracted).IsFalse();
        var amount = result.Fields.Single(_ => _.Path == "Budgets[1].Amount");
        await Assert.That(amount.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(amount.RawText).IsEqualTo("abc");
        await Assert.That(result.Fields.Single(_ => _.Path == "Budgets[0].Amount").State).IsEqualTo(FieldState.Extracted);

        // The row is still applied; the member that failed keeps the element's default.
        var model = NewPlan();
        result.ApplyTo(model);
        await Assert.That(model.Budgets.Count).IsEqualTo(2);
        await Assert.That(model.Budgets[1].Year).IsEqualTo("2027");
        await Assert.That(model.Budgets[1].Amount).IsEqualTo(0m);
    }

    [Test]
    public async Task UnknownRowDropDownValueIsReported()
    {
        using var stream = await Render(
            choreTemplate,
            new ChoreList
            {
                Chores =
                [
                    new()
                    {
                        Name = "Sweep",
                        Priority = Priority.High
                    }
                ]
            });

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            SetText(Control(Items(doc, "Chores")[0], "Priority"), "Urgent");
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<ChoreList>(stream);

        await Assert.That(result.AllExtracted).IsFalse();
        var priority = result.Fields.Single(_ => _.Path == "Chores[0].Priority");
        await Assert.That(priority.State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(priority.RawText).IsEqualTo("Urgent");
    }

    [Test]
    public async Task DeletedRowControlIsReported()
    {
        using var stream = await Render(template, NewPlan());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            Control(Items(doc, "Budgets")[0], "Amount").Remove();
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        await Assert.That(result.AllExtracted).IsFalse();
        var amount = result.Fields.Single(_ => _.Path == "Budgets[0].Amount");
        await Assert.That(amount.State).IsEqualTo(FieldState.Missing);
        await Assert.That(amount.RawText).IsNull();

        var model = NewPlan();
        result.ApplyTo(model);
        await Assert.That(model.Budgets.Count).IsEqualTo(2);
        await Assert.That(model.Budgets[0].Year).IsEqualTo("2026");
        await Assert.That(model.Budgets[0].Amount).IsEqualTo(0m);
    }

    // A row index is the item's position in the rebuilt list, so a dropped row is not counted
    // and reports nothing.
    [Test]
    public async Task RowIndexSkipsClearedRows()
    {
        using var stream = await Render(template, NewPlan());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var items = Items(doc, "Budgets");
            SetText(Control(items[0], "Year"), "");
            SetText(Control(items[0], "Amount"), "");
            SetText(Control(items[1], "Amount"), "abc");
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        await Assert.That(result.Fields.Select(_ => _.Path))
            .IsEquivalentTo(["Budgets[0].Year", "Budgets[0].Amount"]);
        await Assert.That(result.Fields.Single(_ => _.Path == "Budgets[0].Amount").State).IsEqualTo(FieldState.ParseFailed);

        var model = NewPlan();
        result.ApplyTo(model);
        await Assert.That(model.Budgets.Count).IsEqualTo(1);
        await Assert.That(model.Budgets[0].Year).IsEqualTo("2027");
    }

    // Text that does not parse is still something the user typed: the row is kept and reported
    // rather than dropped as cleared, in the blank row of an empty collection too.
    [Test]
    public async Task RowHoldingOnlyUnparseableValueIsKept()
    {
        using var stream = await Render(
            template,
            new BudgetPlan
            {
                Title = "Empty",
                Budgets = []
            });

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            SetText(Control(doc.MainDocumentPart!.Document!.Body!, "Amount"), "abc");
            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<BudgetPlan>(stream);

        await Assert.That(result.AllExtracted).IsFalse();
        await Assert.That(result.Fields.Single(_ => _.Path == "Budgets[0].Amount").State).IsEqualTo(FieldState.ParseFailed);
        await Assert.That(result.Fields.Single(_ => _.Path == "Budgets[0].Year").State).IsEqualTo(FieldState.Empty);

        var model = NewPlan();
        result.ApplyTo(model);
        await Assert.That(model.Budgets.Count).IsEqualTo(1);
    }

    [Test]
    public async Task UntouchedBlankRowReportsNothing()
    {
        using var stream = await Render(
            choreTemplate,
            new ChoreList
            {
                Chores = []
            });

        var result = ParchmentExtractor.Extract<ChoreList>(stream);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(result.Fields.Count).IsEqualTo(0);
    }

    public class Entry
    {
        [EditableField]
        public string? Name { get; set; }

        public string? Code { get; set; }
    }

    [ParchmentBindable]
    public partial class EntryList
    {
        [EditableField]
        public string? Name { get; set; }

        [EditableField]
        public List<Entry> Entries { get; set; } = [];
    }

    const string entryTemplate =
        """
        {% for e in Entries %}

        Name: {{ e.Name }}

        {% endfor %}
        """;

    static EntryList NewEntries() =>
        new()
        {
            Name = "root",
            Entries =
            [
                new()
                {
                    Name = "first",
                    Code = "A"
                },
                new()
                {
                    Name = "second",
                    Code = "B"
                }
            ]
        };

    // Extraction reads one repeating section per collection, so the rows edited in a second one
    // would be discarded.
    [Test]
    [Arguments("Name: {{ other.Name }}")]
    [Arguments("{% if other.Code %}\n\nName: {{ other.Name }}\n\n{% endif %}")]
    public async Task SecondLoopBindingTheCollectionIsRejected(string secondBody)
    {
        using var docx = DocxTemplateBuilder.Build(
            entryTemplate + "\n\nAgain\n\n{% for other in Entries %}\n\n" + secondBody + "\n\n{% endfor %}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EntryList>(docx))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }

    // The member name is the control's tag within a row, and extraction reads the first.
    [Test]
    [Arguments("{{ e.Name }} and {{ e.Name }}")]
    [Arguments("{{ e.Name }}\n\n{% if e.Code %}\n\n{{ e.Name }}\n\n{% endif %}")]
    public async Task MemberBoundTwiceInOneLoopIsRejected(string body)
    {
        using var docx = DocxTemplateBuilder.Build("{% for e in Entries %}\n\n" + body + "\n\n{% endfor %}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<EntryList>(docx))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }

    // A loop that puts no control in its rows has nothing to extract, so it renders as an ordinary
    // read-only loop wherever it sits, and the one loop holding the controls is the one read back.
    [Test]
    public async Task LoopsWithoutControlsRenderReadOnlyBesideTheBoundLoop()
    {
        using var stream = await Render(
            """
            Root: {{ Name }}

            {% for e in Entries %}

            Code {{ e.Code }}

            {% endfor %}

            {% if Entries.size > 0 %}

            Has entries

            {% endif %}

            {% for e in Entries %}

            Name: {{ e.Name }}

            {% endfor %}

            {% for e in Entries %}

            {% for inner in Entries %}

            Pair {{ e.Code }}{{ inner.Name }}

            {% endfor %}

            {% endfor %}
            """,
            NewEntries());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sections = body.Descendants<SdtBlock>()
                .Where(_ => _.SdtProperties?.GetFirstChild<W15.SdtRepeatedSection>() != null)
                .ToList();
            await Assert.That(sections.Count).IsEqualTo(1);

            var text = body.Descendants<Paragraph>().Select(_ => _.InnerText).ToList();
            await Assert.That(text).Contains("Code A");
            await Assert.That(text).Contains("Code B");
            await Assert.That(text).Contains("Has entries");
            await Assert.That(text).Contains("Pair Bfirst");

            var items = sections[0].Descendants<SdtBlock>().Where(IsItem).ToList();
            SetText(Control(items[1], "Name"), "edited");
            doc.Save();
        }

        stream.Position = 0;
        var model = new EntryList();
        var result = ParchmentExtractor.Extract<EntryList>(stream);
        result.ApplyTo(model);

        await Assert.That(result.AllExtracted).IsTrue();
        await Assert.That(model.Name).IsEqualTo("root");
        await Assert.That(string.Join(",", model.Entries.Select(_ => _.Name))).IsEqualTo("first,edited");
    }

    // With no control to read back, the collection is absent from the document rather than empty.
    [Test]
    public async Task LoopWithoutControlsLeavesTheCollectionAlone()
    {
        using var stream = await Render("{% for e in Entries %}\n\nCode {{ e.Code }}\n\n{% endfor %}", NewEntries());

        var model = NewEntries();
        var result = ParchmentExtractor.Extract<EntryList>(stream);
        result.ApplyTo(model);

        await Assert.That(model.Entries.Count).IsEqualTo(2);
    }

    // A document can hold a section no registration produced - one pasted in Word, or rendered by
    // an older version. Its row controls are tagged with the bare member name, which a root member
    // can share.
    [Test]
    public async Task RootFieldDoesNotReadAControlInsideARow()
    {
        using var stream = await Render(entryTemplate, NewEntries());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var section = RepeatingSection(body, "Entries");
            section.InsertAfterSelf(section.CloneNode(true));
            doc.Save();
        }

        stream.Position = 0;
        var model = new EntryList
        {
            Name = "untouched"
        };
        var result = ParchmentExtractor.Extract<EntryList>(stream);
        result.ApplyTo(model);

        await Assert.That(result.Fields.Single(_ => _.Path == "Name").State).IsEqualTo(FieldState.Missing);
        await Assert.That(model.Name).IsEqualTo("untouched");
        await Assert.That(string.Join(",", model.Entries.Select(_ => _.Name))).IsEqualTo("first,second");
    }

    [Test]
    public async Task RowFieldDoesNotReadAControlOfANestedRow()
    {
        using var stream = await Render(entryTemplate, NewEntries());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var section = RepeatingSection(body, "Entries");
            var nested = section.CloneNode(true);
            var first = section.Descendants<SdtBlock>().First(IsItem);
            var control = Control(first, "Name");
            var host = control.Ancestors<Paragraph>().First();
            control.Remove();
            host.InsertAfterSelf(nested);
            doc.Save();
        }

        stream.Position = 0;
        var model = new EntryList();
        ParchmentExtractor.Extract<EntryList>(stream).ApplyTo(model);

        await Assert.That(string.Join(",", model.Entries.Select(_ => _.Name))).IsEqualTo("second");
    }

    const string checklistTemplate =
        """
        {% for i in Items %}

        {{ i.Label }}: {{ i.Done }}

        {% endfor %}
        """;

    static Checklist NewChecklist() =>
        new()
        {
            Items =
            [
                new()
                {
                    Label = "First",
                    Done = true
                },
                new()
                {
                    Label = "Second"
                }
            ]
        };

    [Test]
    public async Task UntouchedAddedRowExtractsNothing()
    {
        using var stream = await Render(checklistTemplate, NewChecklist());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            AddRowAsWord(RepeatingSection(doc.MainDocumentPart!.Document!.Body!, "Items"));
            doc.Save();
        }

        stream.Position = 0;
        var extracted = new Checklist();
        ParchmentExtractor.Extract<Checklist>(stream).ApplyTo(extracted);

        // The added row holds nothing but the placeholders Word reset it to and a box that cannot
        // be anything but unchecked, so it is no entry. The rendered "Second" row, whose box is
        // just as unchecked, is.
        await Assert.That(extracted.Items.Count).IsEqualTo(2);
        await Assert.That(extracted.Items[0].Label).IsEqualTo("First");
        await Assert.That(extracted.Items[1].Label).IsEqualTo("Second");
        await Assert.That(extracted.Items[1].Done).IsFalse();
    }

    [Test]
    public async Task UntouchedAddedRowInArrayExtractsNothing()
    {
        using var stream = await Render(
            checklistTemplate,
            new ChecklistArray
            {
                Items = NewChecklist().Items.ToArray()
            });

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            AddRowAsWord(RepeatingSection(doc.MainDocumentPart!.Document!.Body!, "Items"));
            doc.Save();
        }

        stream.Position = 0;
        var extracted = new ChecklistArray();
        ParchmentExtractor.Extract<ChecklistArray>(stream).ApplyTo(extracted);

        await Assert.That(extracted.Items.Length).IsEqualTo(2);
    }

    [Test]
    public async Task AddedRowWithOnlyTextExtractsOneItem()
    {
        using var stream = await Render(checklistTemplate, NewChecklist());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var added = AddRowAsWord(RepeatingSection(doc.MainDocumentPart!.Document!.Body!, "Items"));
            SetText(Control(added, "Label"), "Third");
            doc.Save();
        }

        stream.Position = 0;
        var extracted = new Checklist();
        ParchmentExtractor.Extract<Checklist>(stream).ApplyTo(extracted);

        await Assert.That(extracted.Items.Count).IsEqualTo(3);
        await Assert.That(extracted.Items[2].Label).IsEqualTo("Third");
        await Assert.That(extracted.Items[2].Done).IsFalse();
    }

    [Test]
    public async Task AddedRowWithOnlyCheckedBoxExtractsOneItem()
    {
        using var stream = await Render(checklistTemplate, NewChecklist());

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var added = AddRowAsWord(RepeatingSection(doc.MainDocumentPart!.Document!.Body!, "Items"));
            Control(added, "Done").SdtProperties!
                .GetFirstChild<W14.SdtContentCheckBox>()!
                .GetFirstChild<W14.Checked>()!
                .Val = W14.OnOffValues.One;
            doc.Save();
        }

        stream.Position = 0;
        var extracted = new Checklist();
        ParchmentExtractor.Extract<Checklist>(stream).ApplyTo(extracted);

        await Assert.That(extracted.Items.Count).IsEqualTo(3);
        await Assert.That(extracted.Items[2].Done).IsTrue();
    }

    [Test]
    public async Task RenderedItemHoldingOnlyFalseRoundTrips()
    {
        // Every control but the box shows a placeholder here too, but it is the one the render
        // wrote, not one Word reset the control to, so the row is a rendered item and is kept.
        using var stream = await Render(
            checklistTemplate,
            new Checklist
            {
                Items =
                [
                    new()
                    {
                        Label = "First",
                        Done = true
                    },
                    new()
                ]
            });

        var extracted = new Checklist();
        ParchmentExtractor.Extract<Checklist>(stream).ApplyTo(extracted);

        await Assert.That(extracted.Items.Count).IsEqualTo(2);
        await Assert.That(extracted.Items[1].Done).IsFalse();
    }

    [ParchmentBindable]
    public partial class ChecklistArray
    {
        [EditableField]
        public ChecklistItem[] Items { get; set; } = [];
    }

    // What Word's "+" does to a repeating section whose rows hold values: the last row is copied,
    // every control in the copy that held a value is reset to a placeholder of five spaces, and
    // every checkbox is unchecked. The copy keeps the tags of the row it came from.
    static SdtBlock AddRowAsWord(SdtBlock container)
    {
        var last = container.Descendants<SdtBlock>().Last(IsItem);
        var added = (SdtBlock)last.CloneNode(true);
        foreach (var control in added.Descendants<SdtElement>())
        {
            var properties = control.SdtProperties!;
            var checkBox = properties.GetFirstChild<W14.SdtContentCheckBox>();
            if (checkBox != null)
            {
                checkBox.GetFirstChild<W14.Checked>()!.Val = W14.OnOffValues.Zero;
                control.Descendants<Text>().Single().Text = "☐";
                continue;
            }

            if (properties.GetFirstChild<ShowingPlaceholder>() != null)
            {
                continue;
            }

            properties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>()!.InsertAfterSelf(new ShowingPlaceholder());
            SetPlaceholderText(control, "     ");
        }

        last.InsertAfterSelf(added);
        return added;
    }

    static void SetPlaceholderText(SdtElement sdt, string text)
    {
        var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
        content.RemoveAllChildren();
        content.AppendChild(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
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
        await Assert.That(unreachable)
            .IsEquivalentTo(
            [
                "Crates[0].Carrier.Contact=Ann",
                "Crates[0].Seal.Colour=red",
                "Crates[1].Carrier.Contact=Ann",
                "Crates[1].Seal.Colour=red"
            ]);

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

    // An item member reached through the loop variable is a control like a root member, and the
    // control is built from the member, not from the expression: a filter would be dropped and the
    // filtered text could not be read back.
    [Test]
    [Arguments("{{ b.Year | append: '!' }}")]
    [Arguments("{{ b.Amount | plus: 1 }}")]
    public async Task FilterOnItemMemberFailsRegistration(string token)
    {
        using var docx = DocxTemplateBuilder.Build(
            $$$"""
               {% for b in Budgets %}

               {{{token}}}

               {% endfor %}
               """);
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<BudgetPlan>(docx))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("plain member-access");
    }

    // The member's path is the control's tag inside each repeated item; extraction reads one.
    [Test]
    public async Task ItemMemberBoundTwiceFailsRegistration()
    {
        using var docx = DocxTemplateBuilder.Build(
            """
            {% for b in Budgets %}

            Year: {{ b.Year }}

            Again: {{ b.Year }}

            {% endfor %}
            """);
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<BudgetPlan>(docx))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }

    // A loop nested in the repeating section renders read-only through Fluid, so the rules for a
    // control do not reach it; neither do they reach the collection's own token.
    [Test]
    public async Task ReadOnlyOccurrencesTakeAFilter()
    {
        using var stream = await Render(
            """
            count={{ Budgets | size }}

            {% for b in Budgets %}

            Year: {{ b.Year }}

            {% for other in Budgets limit:1 %}

            again={{ b.Year | append: '!' }}

            {% endfor %}

            {% endfor %}
            """,
            NewPlan());

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var text = body.Descendants<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.StartsWith("again=", StringComparison.Ordinal) ||
                        _.StartsWith("count=", StringComparison.Ordinal))
            .ToList();
        await Assert.That(text).IsEquivalentTo(["count=2", "again=2026!", "again=2027!"]);
        await Assert.That(body.Descendants<SdtRun>().Count(_ => Tag(_) == "Year")).IsEqualTo(2);
    }

    static bool IsItem(OpenXmlElement element) =>
        element is SdtBlock sdt &&
        sdt.SdtProperties?.GetFirstChild<W15.SdtRepeatedSectionItem>() != null;

    static List<SdtBlock> Items(WordprocessingDocument doc, string tag) =>
        RepeatingSection(doc.MainDocumentPart!.Document!.Body!, tag)
            .ChildElements.First(_ => _.LocalName == "sdtContent")
            .ChildElements.Where(IsItem)
            .Cast<SdtBlock>()
            .ToList();

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
