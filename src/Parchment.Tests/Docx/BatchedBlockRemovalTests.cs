// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// The runner queues the blocks it drops — tag paragraphs, branches that were not chosen, loop
/// bodies — and unlinks them in one pass per parent once there are more than a few. Each test here
/// drops enough to take that pass, and asserts what is left and what was able to see it.
/// </summary>
public partial class BatchedBlockRemovalTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    [ParchmentBindable]
    public partial class Report
    {
        public bool Yes { get; init; } = true;
        public bool No { get; init; }
        public IReadOnlyList<Item> Items { get; init; } = [];
        public MutateToken Mutate { get; init; } = new((_, _) => { });
        public OpenXmlToken Raw { get; init; } = new(_ => []);
    }

    public class Item
    {
        public required string Name { get; init; }
        public bool Active { get; init; }
    }

    [ParchmentBindable]
    public partial class Article
    {
        public bool No { get; init; }

        [Html]
        [EditableField]
        public required string Content { get; set; }
    }

    [Test]
    public async Task ManyConditionalsKeepTheirChosenBranches()
    {
        var blocks = new List<OpenXmlElement>();
        var expected = new List<string>();
        for (var i = 0; i < 40; i++)
        {
            var chosen = i % 3 != 0;
            blocks.Add(BuildParagraph($"before {i}"));
            blocks.Add(BuildParagraph(chosen ? "{% if Yes %}" : "{% if No %}"));
            blocks.Add(BuildParagraph($"then {i}"));
            blocks.Add(BuildParagraph("{% else %}"));
            blocks.Add(BuildParagraph($"otherwise {i}"));
            blocks.Add(BuildParagraph("{% endif %}"));
            expected.Add($"p:before {i}");
            expected.Add(chosen ? $"p:then {i}" : $"p:otherwise {i}");
        }

        using var template = BuildTemplate([.. blocks]);
        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo(string.Join(' ', expected));
    }

    [Test]
    public async Task ElsifChainsAndNestedBlocksKeepTheirOrder()
    {
        var blocks = new List<OpenXmlElement>();
        var expected = new List<string>();
        for (var i = 0; i < 12; i++)
        {
            blocks.AddRange(
                BuildParagraph("{% if No %}"),
                BuildParagraph($"first {i}"),
                BuildParagraph("{% elsif Yes %}"),
                BuildParagraph($"second {i}"),
                BuildParagraph("{% if Yes %}"),
                BuildParagraph($"inner {i}"),
                BuildParagraph("{% if No %}"),
                BuildParagraph($"innermost {i}"),
                BuildParagraph("{% endif %}"),
                BuildParagraph("{% endif %}"),
                BuildParagraph($"third {i}"),
                BuildParagraph("{% else %}"),
                BuildParagraph($"last {i}"),
                BuildParagraph("{% endif %}"));
            expected.AddRange($"p:second {i}", $"p:inner {i}", $"p:third {i}");
        }

        using var template = BuildTemplate([.. blocks]);
        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo(string.Join(' ', expected));
    }

    // Tables that only the dropped blocks kept apart are separated as they are when each block is
    // unlinked on the spot.
    [Test]
    public async Task TablesLeftTogetherByDroppedBlocksAreKeptApart()
    {
        using var template = BuildTemplate(
        [
            .. Hidden(4),
            BuildTable("A"),
            BuildParagraph("{% if Yes %}"),
            BuildParagraph("{% if Yes %}"),
            BuildTable("B"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% if No %}"),
            BuildParagraph("dropped"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% for item in Items %}"),
            BuildParagraph("{% endfor %}"),
            BuildTable("C")
        ]);

        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A p: tbl:B p: tbl:C");
    }

    [Test]
    public async Task TablesAdjacentInTheTemplateStayAsAuthored()
    {
        using var template = BuildTemplate(
        [
            .. Hidden(4),
            BuildParagraph("{% if Yes %}"),
            BuildTable("A"),
            BuildTable("B"),
            BuildParagraph("{% endif %}")
        ]);

        using var doc = await Render(template, new());

        await Assert.That(Blocks(Body(doc))).IsEqualTo("tbl:A tbl:B");
    }

    [Test]
    public async Task LoopBodyDroppingManyBlocksPerIteration()
    {
        var blocks = new List<OpenXmlElement>
        {
            BuildParagraph("{% for item in Items %}")
        };
        for (var i = 0; i < 4; i++)
        {
            blocks.AddRange(
                BuildParagraph("{% if item.Active %}"),
                BuildParagraph($"on {{{{ item.Name }}}} {i}"),
                BuildParagraph("{% else %}"),
                BuildParagraph($"off {{{{ item.Name }}}} {i}"),
                BuildParagraph("{% endif %}"));
        }

        blocks.Add(BuildParagraph("{% endfor %}"));
        blocks.Add(BuildParagraph("end"));

        var items = new List<Item>
        {
            new()
            {
                Name = "a",
                Active = true
            },
            new()
            {
                Name = "b"
            },
            new()
            {
                Name = "c",
                Active = true
            }
        };
        var expected = items
            .SelectMany(item => Enumerable.Range(0, 4).Select(i => $"p:{(item.Active ? "on" : "off")} {item.Name} {i}"))
            .Append("p:end");

        using var template = BuildTemplate([.. Hidden(4), .. blocks]);
        using var doc = await Render(
            template,
            new()
            {
                Items = items
            });

        await Assert.That(Blocks(Body(doc))).IsEqualTo(string.Join(' ', expected));
    }

    [Test]
    public async Task ManyLoopsInOneBody()
    {
        var blocks = new List<OpenXmlElement>();
        var expected = new List<string>();
        for (var i = 0; i < 10; i++)
        {
            blocks.AddRange(
                BuildParagraph($"list {i}"),
                BuildParagraph("{% for item in Items %}"),
                BuildParagraph($"{{{{ item.Name }}}} {i}"),
                BuildParagraph("{% endfor %}"));
            expected.AddRange($"p:list {i}", $"p:a {i}", $"p:b {i}");
        }

        using var template = BuildTemplate([.. blocks]);
        using var doc = await Render(
            template,
            new()
            {
                Items =
                [
                    new()
                    {
                        Name = "a"
                    },
                    new()
                    {
                        Name = "b"
                    }
                ]
            });

        await Assert.That(Blocks(Body(doc))).IsEqualTo(string.Join(' ', expected));
    }

    [Test]
    public async Task ManyBlocksInATableCell()
    {
        var cellBlocks = new List<OpenXmlElement>();
        var expected = new List<string>();
        for (var i = 0; i < 6; i++)
        {
            cellBlocks.AddRange(Hidden(1));
            cellBlocks.Add(BuildParagraph($"kept {i}"));
            expected.Add($"p:kept {i}");
        }

        using var template = BuildTemplate(
        [
            .. Hidden(4),
            BuildParagraph("above"),
            BuildTable([.. cellBlocks]),
            BuildParagraph("below")
        ]);

        using var doc = await Render(template, new());

        var body = Body(doc);
        var cell = body.GetFirstChild<Table>()!.GetFirstChild<TableRow>()!.GetFirstChild<TableCell>()!;
        await Assert.That(Blocks(cell)).IsEqualTo(string.Join(' ', expected));
        await Assert.That(body.ChildElements.Select(_ => _.LocalName)).IsEquivalentTo(["p", "tbl", "p", "sectPr"]);
    }

    [Test]
    public async Task ManyBlocksInAHeader()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(
            [
                BuildParagraph("top"),
                .. Hidden(5),
                BuildParagraph("{% if Yes %}"),
                BuildParagraph("shown"),
                BuildParagraph("{% endif %}"),
                BuildParagraph("bottom")
            ]);
            mainPart.Document = new(
                new Body(
                    BuildParagraph("body"),
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
                        })));
        }

        stream.Position = 0;
        using var rendered = await Render(stream, new());

        var header = rendered.MainDocumentPart!.HeaderParts.Single().Header!;
        await Assert.That(Blocks(header)).IsEqualTo("p:top p:shown p:bottom");
    }

    // A delegate is handed a paragraph of the document, so the blocks dropped before it ran are
    // out of the document by the time it does.
    [Test]
    public async Task MutateDelegateSeesEarlierBlocksGone()
    {
        using var template = BuildTemplate(
        [
            .. Hidden(4),
            BuildParagraph("before"),
            BuildParagraph("{{ Mutate }}"),
            BuildParagraph("after")
        ]);

        string? seen = null;
        using var doc = await Render(
            template,
            new()
            {
                Mutate = new((paragraph, _) => seen = string.Join('|', paragraph.ElementsBefore().Select(_ => _.InnerText)))
            });

        await Assert.That(seen).IsEqualTo("before");
    }

    // Inside a loop the delegate's own paragraph is in the iteration being built, and the rest of
    // the document is still within its reach through the main part.
    [Test]
    public async Task DelegateInALoopSeesEarlierBlocksGoneFromTheBody()
    {
        using var template = BuildTemplate(
        [
            .. Hidden(4),
            BuildParagraph("{% for item in Items %}"),
            BuildParagraph("{% if Yes %}"),
            BuildParagraph("{{ Raw }}"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("{% endfor %}")
        ]);

        var seen = new List<string>();
        using var doc = await Render(
            template,
            new()
            {
                Items =
                [
                    new()
                    {
                        Name = "a"
                    }
                ],
                Raw = new(context =>
                {
                    seen.AddRange(context.MainPart.Document!.Body!.Elements<Paragraph>().Select(_ => _.InnerText));
                    return [BuildParagraph("produced")];
                })
            });

        // The loop's own tags are still there: the body the delegate read is the one being rendered.
        await Assert.That(seen).Contains("{% for item in Items %}");
        await Assert.That(seen).DoesNotContain("hidden");
        await Assert.That(seen).DoesNotContain("{% if No %}");
        await Assert.That(Blocks(Body(doc))).IsEqualTo("p:produced");
    }

    // Whether an editable html control is all its cell holds decides how it is made editable, and
    // a block dropped from the same cell does not count.
    [Test]
    public async Task EditableHtmlIsAloneInItsCellOnceADroppedBlockIsGone()
    {
        using var template = BuildTemplate(
            BuildTable(
            [
                .. Hidden(1),
                BuildParagraph("{{ Content }}")
            ]));
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Article>(template);
        using var stream = new MemoryStream();
        await store.Render(
            new Article
            {
                Content = "<p>one</p><p>two</p>"
            },
            stream);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);

        var row = Body(doc).GetFirstChild<Table>()!.GetFirstChild<TableRow>()!;
        var cell = row.GetFirstChild<TableCell>()!;
        await Assert.That(row.Elements<PermStart>().Count()).IsEqualTo(1);
        await Assert.That(cell.Elements<Paragraph>().Any()).IsFalse();
        await Assert.That(cell.Descendants<PermStart>().Any()).IsFalse();
    }

    [Test]
    public async Task LongBodyRendersTheSameFromManyThreads()
    {
        var blocks = new List<OpenXmlElement>();
        for (var i = 0; i < 60; i++)
        {
            blocks.AddRange(
                BuildParagraph(i % 2 == 0 ? "{% if Yes %}" : "{% if No %}"),
                BuildParagraph($"text {i}"),
                BuildParagraph("{% endif %}"),
                BuildParagraph("{% for item in Items %}"),
                BuildParagraph($"{{{{ item.Name }}}} {i}"),
                BuildParagraph("{% endfor %}"));
        }

        using var template = BuildTemplate([.. blocks]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Report>(template);
        var model = new Report
        {
            Items =
            [
                new()
                {
                    Name = "a"
                },
                new()
                {
                    Name = "b"
                }
            ]
        };

        async Task<byte[]> RenderBytes()
        {
            using var stream = new MemoryStream();
            await store.Render(model, stream);
            return stream.ToArray();
        }

        var expected = await RenderBytes();
        var renders = await Task.WhenAll(Enumerable.Range(0, 32).Select(_ => Task.Run(RenderBytes)));

        foreach (var bytes in renders)
        {
            await Assert.That(bytes.AsSpan().SequenceEqual(expected)).IsTrue();
        }
    }

    // Three blocks dropped per conditional.
    static IEnumerable<OpenXmlElement> Hidden(int count)
    {
        for (var i = 0; i < count; i++)
        {
            yield return BuildParagraph("{% if No %}");
            yield return BuildParagraph("hidden");
            yield return BuildParagraph("{% endif %}");
        }
    }

    static Body Body(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!;

    static string Blocks(OpenXmlElement container) =>
        string.Join(
            ' ',
            container.ChildElements
                .Where(_ => _ is not (SectionProperties or TableCellProperties))
                .Select(_ => $"{_.LocalName}:{_.InnerText}"));

    static async Task<WordprocessingDocument> Render(Stream template, Report model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Report>(template);
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    static MemoryStream BuildTemplate(params OpenXmlElement[] blocks)
    {
        var stream = DocxTemplateBuilder.Build();
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var section = body.GetFirstChild<SectionProperties>()!;
            foreach (var block in blocks)
            {
                body.InsertBefore(block, section);
            }
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    static Table BuildTable(string text) =>
        BuildTable(BuildParagraph(text));

    static Table BuildTable(params OpenXmlElement[] cellBlocks) =>
        new(
            new TableProperties(
                new TableWidth
                {
                    Width = "0",
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(new GridColumn()),
            new TableRow(new TableCell(cellBlocks)));
}
