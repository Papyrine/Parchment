using TUnit.Assertions.Enums;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A footnote or endnote is tied to its reference by id alone, and Word refuses to open a document
/// in which two references name one note — which <c>OpenXmlValidator</c> does not report. A loop
/// clones the reference with its paragraph, so every copy needs a note of its own, and a reference
/// a render removes takes its note with it.
/// </summary>
public partial class NoteReferenceTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [ParchmentBindable]
    public partial class Model
    {
        public string Name { get; init; } = "Alice";
        public bool Flag { get; init; }
        public IReadOnlyList<Line> Lines { get; init; } = [];
        public IReadOnlyList<Group> Groups { get; init; } = [];

        [Html]
        public string Html { get; init; } = "<p>one</p><p>two</p>";
    }

    public class Line
    {
        public required string Name { get; init; }
    }

    public class Group
    {
        public required string Name { get; init; }
        public IReadOnlyList<string> Items { get; init; } = [];
    }

    public class Budget
    {
        [EditableField]
        public required string Year { get; set; }
    }

    [ParchmentBindable]
    public partial class Plan
    {
        [EditableField]
        public required List<Budget> Budgets { get; set; }
    }

    static Model ThreeLines() =>
        new()
        {
            Lines = [new() { Name = "a" }, new() { Name = "b" }, new() { Name = "c" }]
        };

    [Test]
    public async Task LoopGivesEachIterationItsOwnFootnote()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Footnote(1)),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "the note")]);

        using var doc = await Render(template, ThreeLines());

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([1L, 2L, 3L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["1:the note", "2:the note", "3:the note"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    [Test]
    public async Task LoopGivesEachIterationItsOwnEndnote()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Endnote(1)),
                Paragraph("{% endfor %}")
            ],
            endnotes: [EndnoteText(1, "the note")]);

        using var doc = await Render(template, ThreeLines());

        await Assert.That(References<EndnoteReference>(doc)).IsEquivalentTo([1L, 2L, 3L], CollectionOrdering.Matching);
        await Assert.That(Notes<Endnote>(doc)).IsEquivalentTo(["1:the note", "2:the note", "3:the note"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    // Footnotes and endnotes number separately, so a paragraph carrying one of each takes a fresh
    // id from each part, above whatever that part already holds.
    [Test]
    public async Task FootnoteAndEndnoteInOneParagraphAreNumberedApart()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Footnote(7), Endnote(2)),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(7, "foot")],
            endnotes: [EndnoteText(2, "end")]);

        using var doc = await Render(template, ThreeLines());

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([7L, 8L, 9L], CollectionOrdering.Matching);
        await Assert.That(References<EndnoteReference>(doc)).IsEquivalentTo([2L, 3L, 4L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["7:foot", "8:foot", "9:foot"], CollectionOrdering.Matching);
        await Assert.That(Notes<Endnote>(doc)).IsEquivalentTo(["2:end", "3:end", "4:end"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    [Test]
    public async Task NestedLoopsGiveEveryReferenceItsOwnNote()
    {
        using var template = Build(
            [
                Paragraph("{% for group in Groups %}"),
                Paragraph("{{ group.Name }}", Footnote(1)),
                Paragraph("{% for item in group.Items %}"),
                Paragraph("{{ item }}", Footnote(2)),
                Paragraph("{% endfor %}"),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "group note"), FootnoteText(2, "item note")]);

        using var doc = await Render(
            template,
            new()
            {
                Groups =
                [
                    new()
                    {
                        Name = "A",
                        Items = ["x", "y"]
                    },
                    new()
                    {
                        Name = "B",
                        Items = ["z"]
                    }
                ]
            });

        // In reading order: A, x, y, B, z. The first reference to each note keeps it.
        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([1L, 2L, 3L, 4L, 5L], CollectionOrdering.Matching);
        // The copies of one note sit together in the part, after the note they were made from.
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(
            [
                "1:group note",
                "4:group note",
                "2:item note",
                "3:item note",
                "5:item note"
            ],
            CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    [Test]
    public async Task ReferenceInsideATableInsideALoopIsCovered()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                new Table(
                    new TableProperties(
                        new TableWidth
                        {
                            Type = TableWidthUnitValues.Auto
                        }),
                    new TableGrid(new GridColumn()),
                    new TableRow(new TableCell(Paragraph("{{ line.Name }}", Footnote(1), Endnote(1))))),
                Paragraph("between"),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "foot")],
            endnotes: [EndnoteText(1, "end")]);

        using var doc = await Render(template, ThreeLines());

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([1L, 2L, 3L], CollectionOrdering.Matching);
        await Assert.That(References<EndnoteReference>(doc)).IsEquivalentTo([1L, 2L, 3L], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    // The loop over an [EditableField] collection is the other place a loop body is cloned.
    [Test]
    public async Task RepeatingSectionGivesEachItemItsOwnFootnote()
    {
        using var template = Build(
            [
                Paragraph("{% for budget in Budgets %}"),
                Paragraph("{{ budget.Year }}", Footnote(1)),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "the note")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Plan>(template);

        using var doc = await Render(
            store,
            new Plan
            {
                Budgets = [new() { Year = "2026" }, new() { Year = "2027" }]
            });

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([1L, 2L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["1:the note", "2:the note"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    [Test]
    public async Task EmptyLoopTakesItsNotes()
    {
        using var template = Build(
            [
                Paragraph("kept", Footnote(2)),
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Footnote(1), Endnote(1)),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "looped"), FootnoteText(2, "kept")],
            endnotes: [EndnoteText(1, "looped")]);

        using var doc = await Render(template, new());

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([2L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["2:kept"], CollectionOrdering.Matching);
        await Assert.That(References<EndnoteReference>(doc)).IsEmpty();
        await Assert.That(Notes<Endnote>(doc)).IsEmpty();
        await AssertPaired(doc);
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task UnchosenBranchTakesItsNote(bool flag)
    {
        using var template = Build(
            [
                Paragraph("{% if Flag %}"),
                Paragraph("yes", Footnote(1), Endnote(1)),
                Paragraph("{% else %}"),
                Paragraph("no", Footnote(2), Endnote(2)),
                Paragraph("{% endif %}")
            ],
            footnotes: [FootnoteText(1, "yes foot"), FootnoteText(2, "no foot")],
            endnotes: [EndnoteText(1, "yes end"), EndnoteText(2, "no end")]);

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        var (id, name) = flag ? (1L, "yes") : (2L, "no");
        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([id], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo([$"{id}:{name} foot"], CollectionOrdering.Matching);
        await Assert.That(References<EndnoteReference>(doc)).IsEquivalentTo([id], CollectionOrdering.Matching);
        await Assert.That(Notes<Endnote>(doc)).IsEquivalentTo([$"{id}:{name} end"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    // A branch removed inside a loop: the iterations that keep it each get a note, and the note
    // goes only when no iteration kept it.
    [Test]
    public async Task ConditionalInsideALoopCountsTheReferencesThatSurvive()
    {
        using var template = Build(
            [
                Paragraph("{% for group in Groups %}"),
                Paragraph("{% if group.Items.size > 0 %}"),
                Paragraph("{{ group.Name }}", Footnote(1)),
                Paragraph("{% else %}"),
                Paragraph("{{ group.Name }} is empty", Footnote(2)),
                Paragraph("{% endif %}"),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "has items"), FootnoteText(2, "has none")]);

        using var doc = await Render(
            template,
            new()
            {
                Groups =
                [
                    new()
                    {
                        Name = "A",
                        Items = ["x"]
                    },
                    new()
                    {
                        Name = "B",
                        Items = ["y"]
                    }
                ]
            });

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([1L, 3L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["1:has items", "3:has items"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    // A note binds against the model, in its own part, so its tokens are substituted once and each
    // iteration's copy carries the result. The loop variable is not in scope there.
    [Test]
    public async Task CopiedNoteCarriesItsRenderedText()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Footnote(1)),
                Paragraph("{% endfor %}")
            ],
            footnotes:
            [
                new(
                    Paragraph("{% if Flag %}"),
                    Paragraph("hidden"),
                    Paragraph("{% else %}"),
                    new Paragraph(new Run(new FootnoteReferenceMark()), Run("note for {{ Name }}")),
                    Paragraph("{% endif %}"))
                {
                    Id = 1
                }
            ]);

        using var doc = await Render(template, ThreeLines());

        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(
            [
                "1:note for Alice",
                "2:note for Alice",
                "3:note for Alice"
            ],
            CollectionOrdering.Matching);
        var footnotes = doc.MainDocumentPart!.FootnotesPart!.Footnotes!;
        await Assert.That(footnotes.Descendants<BookmarkStart>().Count()).IsEqualTo(0);
        await Assert.That(footnotes.Descendants<BookmarkEnd>().Count()).IsEqualTo(0);
        await AssertPaired(doc);
    }

    [Test]
    public async Task LoopVariableInANoteIsRejectedAtRegistration()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Footnote(1)),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "note for {{ line.Name }}")]);

        var store = new TemplateStore();
        await Assert.That(() => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
    }

    // Nothing here is repeated or removed, so the notes keep their ids — including a note the
    // template never referenced, which is the template's own business.
    [Test]
    public async Task ReferencesARenderLeavesAloneKeepTheirNotes()
    {
        using var template = Build(
            [
                Paragraph("{{ Name }}", Footnote(4)),
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}"),
                Paragraph("{% endfor %}"),
                Paragraph("static", Endnote(3))
            ],
            footnotes: [FootnoteText(4, "referenced"), FootnoteText(9, "never referenced")],
            endnotes: [EndnoteText(3, "referenced")]);

        using var doc = await Render(template, ThreeLines());

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([4L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["4:referenced", "9:never referenced"], CollectionOrdering.Matching);
        await Assert.That(References<EndnoteReference>(doc)).IsEquivalentTo([3L], CollectionOrdering.Matching);
        await Assert.That(Notes<Endnote>(doc)).IsEquivalentTo(["3:referenced"], CollectionOrdering.Matching);
    }

    // A block-shaped or inline value in a paragraph that also holds a reference splits or splices
    // that paragraph. Each reference lands in one half only, so there is nothing to copy.
    [Test]
    [Arguments("<p>one</p><p>two</p>")]
    [Arguments("<b>bold</b>")]
    public async Task SplicedParagraphKeepsEachReferenceOnce(string html)
    {
        using var template = Build(
            [
                new Paragraph(
                    new Run(new FootnoteReference { Id = 1 }),
                    Run("before {{ Html }} after"),
                    new Run(new FootnoteReference { Id = 2 }))
            ],
            footnotes: [FootnoteText(1, "first"), FootnoteText(2, "second")]);

        using var doc = await Render(
            template,
            new()
            {
                Html = html
            });

        await Assert.That(References<FootnoteReference>(doc)).IsEquivalentTo([1L, 2L], CollectionOrdering.Matching);
        await Assert.That(Notes<Footnote>(doc)).IsEquivalentTo(["1:first", "2:second"], CollectionOrdering.Matching);
        await AssertPaired(doc);
    }

    [Test]
    public async Task NoteIdsAreTheSameOnEveryRender()
    {
        using var template = Build(
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}", Footnote(1), Endnote(1)),
                Paragraph("{% endfor %}")
            ],
            footnotes: [FootnoteText(1, "foot")],
            endnotes: [EndnoteText(1, "end")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var first = new MemoryStream();
        await store.Render(ThreeLines(), first);
        using var second = new MemoryStream();
        await store.Render(ThreeLines(), second);

        await Assert.That(second.ToArray()).IsEquivalentTo(first.ToArray(), CollectionOrdering.Matching);
    }

    static async Task AssertPaired(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart!;

        var footnoteReferences = References<FootnoteReference>(doc);
        await Assert.That(footnoteReferences.Distinct().Count()).IsEqualTo(footnoteReferences.Count);
        await Assert.That(NoteIds<Footnote>(mainPart.FootnotesPart?.Footnotes)).IsEquivalentTo(footnoteReferences);

        var endnoteReferences = References<EndnoteReference>(doc);
        await Assert.That(endnoteReferences.Distinct().Count()).IsEqualTo(endnoteReferences.Count);
        await Assert.That(NoteIds<Endnote>(mainPart.EndnotesPart?.Endnotes)).IsEquivalentTo(endnoteReferences);

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    /// <summary>
    /// The ids the body refers to, in reading order.
    /// </summary>
    static List<long> References<T>(WordprocessingDocument doc)
        where T : FootnoteEndnoteReferenceType =>
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<T>()
            .Select(_ => _.Id!.Value)
            .ToList();

    /// <summary>
    /// Every note that is not a separator, as <c>id:text</c>, in the order the part holds them.
    /// </summary>
    static List<string> Notes<T>(WordprocessingDocument doc)
        where T : FootnoteEndnoteType
    {
        var mainPart = doc.MainDocumentPart!;
        OpenXmlPartRootElement? root = mainPart.FootnotesPart?.Footnotes;
        if (typeof(T) == typeof(Endnote))
        {
            root = mainPart.EndnotesPart?.Endnotes;
        }

        return Content<T>(root)
            .Select(_ => $"{_.Id!.Value}:{_.InnerText}")
            .ToList();
    }

    static List<long> NoteIds<T>(OpenXmlPartRootElement? root)
        where T : FootnoteEndnoteType =>
        Content<T>(root)
            .Select(_ => _.Id!.Value)
            .Order()
            .ToList();

    static IEnumerable<T> Content<T>(OpenXmlPartRootElement? root)
        where T : FootnoteEndnoteType
    {
        if (root == null)
        {
            return [];
        }

        return root.Elements<T>().Where(_ => _.Type == null);
    }

    static Task<WordprocessingDocument> Render(Stream template, Model model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        return Render(store, model);
    }

    static async Task<WordprocessingDocument> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    /// <summary>
    /// A template whose body is <paramref name="body"/>, with a notes part for each list given.
    /// Each part opens with the two separator notes Word writes, so the content notes are not the
    /// only ones in it.
    /// </summary>
    static MemoryStream Build(OpenXmlElement[] body, Footnote[]? footnotes = null, Endnote[]? endnotes = null)
    {
        var stream = DocxTemplateBuilder.Build();
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var section = mainPart.Document!.Body!.GetFirstChild<SectionProperties>()!;
            foreach (var element in body)
            {
                section.InsertBeforeSelf(element);
            }

            if (footnotes != null)
            {
                var part = mainPart.AddNewPart<FootnotesPart>("rIdFootnotes");
                part.Footnotes = new(
                    new Footnote(new Paragraph(new Run(new SeparatorMark())))
                    {
                        Type = FootnoteEndnoteValues.Separator,
                        Id = -1
                    },
                    new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
                    {
                        Type = FootnoteEndnoteValues.ContinuationSeparator,
                        Id = 0
                    });
                part.Footnotes.Append(footnotes);
            }

            if (endnotes != null)
            {
                var part = mainPart.AddNewPart<EndnotesPart>("rIdEndnotes");
                part.Endnotes = new(
                    new Endnote(new Paragraph(new Run(new SeparatorMark())))
                    {
                        Type = FootnoteEndnoteValues.Separator,
                        Id = -1
                    },
                    new Endnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
                    {
                        Type = FootnoteEndnoteValues.ContinuationSeparator,
                        Id = 0
                    });
                part.Endnotes.Append(endnotes);
            }
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph Paragraph(string text, params OpenXmlElement[] references) =>
        new([Run(text), ..references]);

    static Run Run(string text) =>
        new(
            new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

    static Run Footnote(long id) =>
        new(
            new FootnoteReference
            {
                Id = id
            });

    static Run Endnote(long id) =>
        new(
            new EndnoteReference
            {
                Id = id
            });

    static Footnote FootnoteText(long id, string text) =>
        new(new Paragraph(new Run(new FootnoteReferenceMark()), Run(text)))
        {
            Id = id
        };

    static Endnote EndnoteText(long id, string text) =>
        new(new Paragraph(new Run(new EndnoteReferenceMark()), Run(text)))
        {
            Id = id
        };
}
