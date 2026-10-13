using DocumentFormat.OpenXml.Validation;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A bookmark's <c>w:id</c> has to be unique in the document. Everything that produces content
/// during a render — the html converter, the markdown renderer — numbers its own bookmarks from
/// 1, so each fragment is renumbered as it joins the document.
/// </summary>
public partial class BookmarkIdTests
{
    [AttributeUsage(AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        public required string Title;

        [Html]
        public required string Body;

        [Markdown]
        public required string Notes;
    }

    static Doc Model() =>
        new()
        {
            Title = "Report",
            Body = """<p id="x">hi</p><p id="y">yo</p>""",
            Notes = "## Notes {#notes}"
        };

    // The markdown is numbered when it is rendered, and the html is converted afterwards, when its
    // marker is swapped - so the converter's 1 and 2 used to land on the headings' 1 and 2.
    [Test]
    public async Task MarkdownFlowNumbersParkedHtmlAfterTheMarkdown()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("# Title {#top}\n\n{{ Body }}\n\n## Other {#other}\n");

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["top", "x", "y", "other"]);
        await AssertUniqueAndPaired(doc);
    }

    // Each conversion used to restart at 1 against the ids the template already had.
    [Test]
    public async Task DocxFlowNumbersProducedBookmarksAboveTheTemplates()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Body }}"),
                Paragraph("{{ Notes }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["existing", "x", "y", "notes"]);
        await AssertUniqueAndPaired(doc);

        // The template's own bookmark is left exactly as authored.
        var existing = Starts(doc).Single(_ => _.Name?.Value == "existing");
        await Assert.That(existing.Id?.Value).IsEqualTo("1");
    }

    // The anchors that mark token paragraphs are bookmarks too, numbered from the template's
    // highest id, and are stripped before the save by pairing an end to an anchor's start by id. A
    // produced bookmark that shared an anchor's id lost its end to that pass.
    [Test]
    public async Task DocxFlowProducedBookmarkKeepsItsEndBesideAnchoredParagraphs()
    {
        using var template = BuildDocx(
            [
                Paragraph("{{ Title }}"),
                Paragraph("{{ Body }}"),
                Paragraph("{{ Title }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["x", "y"]);
        await AssertUniqueAndPaired(doc);
    }

    // Ids are unique across the document rather than within a part, so what is produced into a
    // header is numbered from the same counter as the body.
    [Test]
    public async Task DocxFlowIdsAreUniqueAcrossParts()
    {
        using var template = BuildDocx(
            [Paragraph("{{ Body }}")],
            [
                Bookmarked("header bookmark", "inHeader", "7"),
                Paragraph("{{ Body }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["x", "y", "inHeader", "x", "y"]);
        await AssertUniqueAndPaired(doc);
    }

    // A header with no token is never read by a render, and its bookmark still has to be stepped
    // past: the floor comes from every part at registration, not from the parts a render opens.
    [Test]
    public async Task DocxFlowNumbersAboveATokenFreeHeader()
    {
        using var template = BuildDocx(
            [Paragraph("{{ Body }}")],
            [Bookmarked("header bookmark", "inHeader", "7")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["x", "y", "inHeader"]);
        await AssertUniqueAndPaired(doc);
        var produced = Starts(doc)
            .Where(_ => _.Name?.Value != "inHeader")
            .Select(_ => int.Parse(_.Id!.Value!));
        await Assert.That(produced.Min()).IsGreaterThan(7);
    }

    // The markdown replaces the body, but a header arrives from the style source as authored.
    [Test]
    public async Task MarkdownFlowNumbersAboveTheStyleSourcesHeader()
    {
        using var styleSource = BuildDocx(
            [Bookmarked("replaced by the markdown", "gone", "9")],
            [Bookmarked("header bookmark", "inHeader", "1")]);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("# Title {#top}\n\n{{ Body }}\n", styleSource);

        using var stream = await Render(store, Model());
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["top", "x", "y", "inHeader"]);
        await AssertUniqueAndPaired(doc);
    }

    [Test]
    public async Task IdsAreTheSameOnEveryRender()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Body }}"),
                Paragraph("{{ Notes }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var first = await Render(store, Model());
        using var second = await Render(store, Model());
        using var firstDoc = WordprocessingDocument.Open(first, false);
        using var secondDoc = WordprocessingDocument.Open(second, false);

        await Assert.That(Ids(secondDoc)).IsEquivalentTo(Ids(firstDoc));
    }

    [ParchmentBindable]
    public partial class Marked
    {
        public required string Title;
        public MutateToken Mark = new((_, _) => { });
        public OpenXmlToken Raw = OpenXmlToken.Empty;
        public IReadOnlyList<MarkedLine> Lines = [];
    }

    public class MarkedLine
    {
        public MutateToken Mark { get; init; } = new((_, _) => { });
    }

    // What a delegate can do: it cannot see which ids the document has taken, so it picks one.
    static MutateToken Mark(string name, string id = "1") =>
        new((paragraph, _) =>
            paragraph.Append(
                new BookmarkStart
                {
                    Id = id,
                    Name = name
                },
                new Run(new Text("marked")),
                new BookmarkEnd
                {
                    Id = id
                }));

    // A MutateToken edits its host in place, so nothing renumbered what it added, and the id it
    // picked was an anchor's: the strip took the end and Word listed no bookmark at all.
    [Test]
    public async Task MutateTokenBookmarkKeepsItsEnd()
    {
        using var template = BuildDocx(
            [
                Paragraph("{{ Title }}"),
                Paragraph("{{ Mark }}"),
                Paragraph("{{ Title }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        foreach (var id in new[] { "1", "2", "3", "40" })
        {
            using var stream = await Render(
                store,
                new Marked
                {
                    Title = "Report",
                    Mark = Mark("Injected", id)
                });
            using var doc = WordprocessingDocument.Open(stream, false);

            await Assert.That(Names(doc)).IsEquivalentTo(["Injected"]);
            await AssertUniqueAndPaired(doc);
        }
    }

    // The host carries a bookmark of the template's with the very id the delegate picks. Only what
    // the delegate added is renumbered.
    [Test]
    public async Task MutateTokenBookmarkIsNumberedAboveTheTemplates()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Bookmarked("{{ Mark }}", "host", "2")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Mark = Mark("Injected", "2")
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["existing", "host", "Injected"]);
        await AssertUniqueAndPaired(doc);

        var starts = Starts(doc).ToDictionary(_ => _.Name!.Value!, _ => _.Id!.Value!);
        await Assert.That(starts["existing"]).IsEqualTo("1");
        await Assert.That(starts["host"]).IsEqualTo("2");
        await Assert.That(int.Parse(starts["Injected"])).IsGreaterThan(2);
    }

    // One delegate per iteration, each writing the same id into a clone of the same paragraph.
    [Test]
    public async Task MutateTokenBookmarkInALoopIsNumberedPerIteration()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Mark }}"),
                Paragraph("{% endfor %}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Lines =
                [
                    new()
                    {
                        Mark = Mark("First")
                    },
                    new()
                    {
                        Mark = Mark("Second")
                    },
                    new()
                    {
                        Mark = Mark("Third")
                    }
                ]
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["existing", "First", "Second", "Third"]);
        await AssertUniqueAndPaired(doc);
    }

    // A header's anchors are numbered within the header, so the id a delegate picks there is as
    // likely to be one of theirs, and what it adds still has to clear the body's bookmarks.
    [Test]
    public async Task MutateTokenBookmarkInAHeaderIsNumberedWithTheBody()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Raw }}")
            ],
            [
                Paragraph("{{ Title }}"),
                Paragraph("{{ Mark }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Mark = Mark("Injected"),
                Raw = Produced()
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["existing", "Produced1", "Injected"]);
        await AssertUniqueAndPaired(doc);
    }

    // Nested containers go through the same runner and the same strip: a paragraph in a table
    // cell, inside a branch that is kept.
    [Test]
    public async Task MutateTokenBookmarkInATableCellKeepsItsEnd()
    {
        using var template = BuildDocx(
            [
                Paragraph("{% if Title %}"),
                new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn()),
                    new TableRow(
                        new TableCell(
                            Paragraph("{{ Title }}"),
                            Paragraph("{{ Mark }}")))),
                Paragraph("{% endif %}"),
                Paragraph("after")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Mark = Mark("Injected", "2")
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["Injected"]);
        await AssertUniqueAndPaired(doc);
    }

    // A markdown template has no host paragraph in its body, but its style source's header is a
    // docx part and binds through the same runner.
    [Test]
    public async Task MarkdownFlowMutateTokenBookmarkInAHeaderKeepsItsEnd()
    {
        using var styleSource = BuildDocx(
            [Paragraph("replaced by the markdown")],
            [
                Bookmarked("header bookmark", "inHeader", "1"),
                Paragraph("{{ Title }}"),
                Paragraph("{{ Mark }}")
            ]);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Marked>("# Title {#top}\n\nText\n", styleSource);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Mark = Mark("Injected", "2")
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["top", "inHeader", "Injected"]);
        await AssertUniqueAndPaired(doc);
    }

    // The strip takes an anchor's own end and no other, whatever id a bookmark elsewhere in the
    // part carries. Here the delegate adds a paragraph beside its host, which is not renumbered.
    [Test]
    public async Task BookmarkBesideTheHostKeepsItsEnd()
    {
        using var template = BuildDocx(
            [
                Paragraph("{{ Title }}"),
                Paragraph("{{ Mark }}"),
                Paragraph("{{ Title }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        foreach (var id in new[] { "1", "2", "3" })
        {
            using var stream = await Render(
                store,
                new Marked
                {
                    Title = "Report",
                    Mark = new((paragraph, _) => paragraph.InsertAfterSelf(Bookmarked("beside", "Beside", id)))
                });
            using var doc = WordprocessingDocument.Open(stream, false);

            await Assert.That(Names(doc)).IsEquivalentTo(["Beside"]);
            await AssertUniqueAndPaired(doc);
        }
    }

    // A bookmark that opens in the host and closes outside it is left on the id it was given:
    // renumbering the half that can be seen would part it from the half that cannot.
    [Test]
    public async Task BookmarkLeavingTheHostStaysPaired()
    {
        using var template = BuildDocx(
            [
                Paragraph("{{ Title }}"),
                Paragraph("{{ Mark }}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Mark = new((paragraph, _) =>
                {
                    paragraph.Append(
                        new BookmarkStart
                        {
                            Id = "1",
                            Name = "Spanning"
                        });
                    paragraph.InsertAfterSelf(
                        new Paragraph(
                            new Run(new Text("next")),
                            new BookmarkEnd
                            {
                                Id = "1"
                            }));
                })
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["Spanning"]);
        await AssertUniqueAndPaired(doc);
    }

    // An OpenXmlToken's elements are a fragment, renumbered as they are produced: alone in the
    // paragraph, spliced into one, and once per loop iteration.
    [Test]
    public async Task OpenXmlTokenBookmarksAreNumberedAboveTheTemplates()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Raw }}"),
                Paragraph("see {{ Raw }} here"),
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ Raw }}"),
                Paragraph("{% endfor %}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        using var stream = await Render(
            store,
            new Marked
            {
                Title = "Report",
                Raw = Produced(),
                Lines = [new(), new()]
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Names(doc)).IsEquivalentTo(["existing", "Produced1", "Produced2", "Produced3", "Produced4"]);
        await AssertUniqueAndPaired(doc);
        await Assert.That(Starts(doc).Single(_ => _.Name?.Value == "existing").Id?.Value).IsEqualTo("1");
    }

    [Test]
    public async Task MutateTokenIdsAreTheSameOnEveryRender()
    {
        using var template = BuildDocx(
            [
                Bookmarked("has bookmark", "existing", "1"),
                Paragraph("{{ Mark }}"),
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Mark }}"),
                Paragraph("{% endfor %}")
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Marked>(template);

        static Marked Build() =>
            new()
            {
                Title = "Report",
                Mark = Mark("Injected"),
                Lines =
                [
                    new()
                    {
                        Mark = Mark("First")
                    },
                    new()
                    {
                        Mark = Mark("Second")
                    }
                ]
            };

        using var first = await Render(store, Build());
        using var second = await Render(store, Build());
        using var firstDoc = WordprocessingDocument.Open(first, false);
        using var secondDoc = WordprocessingDocument.Open(second, false);

        await AssertUniqueAndPaired(firstDoc);
        await Assert.That(Ids(secondDoc)).IsEquivalentTo(Ids(firstDoc));
    }

    // The same delegate runs once per site, and a bookmark's name has to be unique as well.
    static OpenXmlToken Produced()
    {
        var count = 0;
        return new(_ => [Bookmarked("produced", $"Produced{++count}", "1")]);
    }

    static async Task AssertUniqueAndPaired(WordprocessingDocument doc)
    {
        var starts = Ids(doc);
        await Assert.That(starts.Distinct().Count()).IsEqualTo(starts.Count);

        // Every start is closed, once, in the part it opens in.
        foreach (var root in Roots(doc))
        {
            var opened = root.Descendants<BookmarkStart>().Select(_ => _.Id?.Value ?? "").Order().ToList();
            var closed = root.Descendants<BookmarkEnd>().Select(_ => _.Id?.Value ?? "").Order().ToList();
            await Assert.That(closed).IsEquivalentTo(opened);
        }

        var errors = new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static List<BookmarkStart> Starts(WordprocessingDocument doc) =>
        Roots(doc).SelectMany(_ => _.Descendants<BookmarkStart>()).ToList();

    static List<string> Names(WordprocessingDocument doc) =>
        Starts(doc).Select(_ => _.Name?.Value ?? "").ToList();

    static List<string> Ids(WordprocessingDocument doc) =>
        Starts(doc).Select(_ => _.Id?.Value ?? "").ToList();

    static IEnumerable<OpenXmlElement> Roots(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart!;
        yield return mainPart.Document!.Body!;
        foreach (var header in mainPart.HeaderParts)
        {
            yield return header.Header!;
        }
    }

    static async Task<MemoryStream> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildDocx(OpenXmlElement[] body, Paragraph[]? header = null)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var documentBody = mainPart.Document.Body!;
            documentBody.Append(body);

            var section = new SectionProperties();
            if (header != null)
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
                headerPart.Header = new(header);
                section.Append(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    });
            }

            section.Append(
                new PageSize
                {
                    Width = 6500,
                    Height = 8000
                });
            documentBody.Append(section);
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph Paragraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    static Paragraph Bookmarked(string text, string name, string id) =>
        new(
            new BookmarkStart
            {
                Id = id,
                Name = name
            },
            new Run(new Text(text)),
            new BookmarkEnd
            {
                Id = id
            });
}
