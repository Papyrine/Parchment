// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A registered docx template parses its token-bearing parts once and every render clones them.
/// The tree is shared by concurrent renders, so no render may change it, and a clone has to be the
/// tree a parse of the snapshot gives, in every part a token can sit in.
/// </summary>
public partial class ParsedPartsTests
{
    [ParchmentBindable]
    public partial class Report
    {
        public string Title { get; init; } = "Report";
        public bool Flag { get; init; } = true;
        public string Notes { get; init; } = "";
        public IReadOnlyList<Row> Rows { get; init; } = [];

        // Edits its host paragraph in place, which has to be this render's own copy.
        public MutateToken Mark { get; init; } = new((paragraph, _) => paragraph.Append(new Run(new Text("marked"))));
    }

    public class Row
    {
        public required string Name { get; init; }
        public required int Amount { get; init; }
    }

    public class Budget
    {
        [EditableField]
        public required string Year { get; set; }

        [EditableField]
        public required decimal Amount { get; set; }
    }

    [ParchmentBindable]
    public partial class Order
    {
        public required string Number;

        [EditableField]
        public required string PurchaseOrder { get; set; }

        [EditableField]
        public bool Approved { get; set; }

        [EditableField]
        public string? Notes { get; set; }

        [EditableField]
        public required List<Budget> Budgets { get; set; }
    }

    const int threads = 16;
    const int rendersPerThread = 12;

    static Report[] Reports() =>
    [
        new(),
        new()
        {
            Title = "Second <&> report",
            Flag = false,
            Notes = "A **bold** note\n\n- one\n- two",
            Rows = [new() { Name = "a", Amount = 1 }, new() { Name = "b", Amount = 2 }]
        },
        new()
        {
            Title = "Line one\nline two",
            Notes = "Plain",
            Rows = Enumerable.Range(0, 20).Select(_ => new Row { Name = "row " + _, Amount = _ }).ToList()
        },
        new()
        {
            Title = "",
            Flag = false
        }
    ];

    [Test]
    public async Task ConcurrentRendersMatchSequentialOnesInEveryPart()
    {
        var models = Reports();
        var expected = new byte[models.Length][];
        for (var index = 0; index < models.Length; index++)
        {
            // A store of its own per model, so each expectation is the first render of a template
            // nothing else has rendered.
            expected[index] = await Render(ReportStore(), models[index]);
        }

        // Every part bound, so the comparison below is over substituted content in all five.
        foreach (var entry in new[] { "word/document.xml", "word/header1.xml", "word/footer1.xml", "word/footnotes.xml", "word/endnotes.xml" })
        {
            var text = Text(expected[1], entry);
            await Assert.That(text).Contains("Second &lt;&amp;&gt; report");
            await Assert.That(text).DoesNotContain("{{");
            await Assert.That(text).DoesNotContain(Anchors.Prefix);
        }

        // The first render of this store is among the concurrent ones, so the threads also race
        // for the one parse.
        var store = ReportStore();
        var results = RenderConcurrently(index => Render(store, models[index % models.Length]));

        for (var thread = 0; thread < threads; thread++)
        {
            foreach (var result in results[thread])
            {
                await Assert.That(result.AsSpan().SequenceEqual(expected[thread % models.Length])).IsTrue();
            }
        }

        // And nothing the concurrent renders did is left behind in the shared tree.
        for (var index = 0; index < models.Length; index++)
        {
            await Assert.That(await Render(store, models[index])).IsEquivalentTo(expected[index]);
        }
    }

    [Test]
    public async Task RenderedDocumentIsValid()
    {
        var store = ReportStore();
        await Render(store, Reports()[2]);

        // A later render, built from a clone of a tree an earlier render was also cloned from.
        var output = await Render(store, Reports()[1]);

        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        // The schema the validator holds has no w:smartTag, which is the template's own.
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Where(_ => !_.Description.Contains("smartTag"));
        await Assert.That(errors).IsEmpty();
        var body =doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).Contains("Second <&> report in a smart tag");
        await Assert.That(body.InnerText).Contains("marked");
        await Assert.That(body.InnerText).Contains("Not flagged");
        await Assert.That(body.InnerText).Contains("b: 2");
        await Assert.That(body.Descendants<TableCell>().Single().InnerText).IsEqualTo("Cell Second <&> report");
    }

    [Test]
    public async Task ConcurrentEditableRendersMatchAndExtract()
    {
        Order[] models =
        [
            NewOrder("ORD-1", "PO-1", true, null, 2),
            NewOrder("ORD-2", "PO-2", false, "call first", 0),
            NewOrder("ORD-3", "", true, "x", 5)
        ];
        var expected = new byte[models.Length][];
        for (var index = 0; index < models.Length; index++)
        {
            expected[index] = await Render(OrderStore(), models[index]);
        }

        var store = OrderStore();
        var results = RenderConcurrently(index => Render(store, models[index % models.Length]));

        for (var thread = 0; thread < threads; thread++)
        {
            foreach (var result in results[thread])
            {
                await Assert.That(result.AsSpan().SequenceEqual(expected[thread % models.Length])).IsTrue();
            }
        }

        var output = results[threads - 1][rendersPerThread - 1];
        var rendered = models[(threads - 1) % models.Length];
        using (var doc = WordprocessingDocument.Open(new MemoryStream(output), false))
        {
            await Assert.That(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc)).IsEmpty();
        }

        var extracted = ParchmentExtractor.Extract<Order>(new MemoryStream(output));
        await Assert.That(extracted.AllExtracted).IsTrue();
        var target = NewOrder("ORD-1", "other", !rendered.Approved, "other", 1);
        extracted.ApplyTo(target);
        await Assert.That(target.PurchaseOrder).IsEqualTo(rendered.PurchaseOrder);
        await Assert.That(target.Approved).IsEqualTo(rendered.Approved);
        await Assert.That(target.Budgets.Select(_ => _.Year)).IsEquivalentTo(rendered.Budgets.Select(_ => _.Year));
    }

    // Word writes namespace declarations, markup-compatibility attributes and attribute order the
    // SDK does not, which is where a clone and a parse could come apart.
    [Test]
    public async Task ConcurrentRendersOfAWordAuthoredTemplateMatch()
    {
        var bytes = await File.ReadAllBytesAsync(Path.Combine(ProjectFiles.ProjectDirectory, "Docx", "WordAuthored.docx"));
        TokenFreePartTests.Model[] models =
        [
            new(),
            new()
            {
                Name = "Initech",
                Flag = false,
                Lines = [new() { Name = "first" }, new() { Name = "second" }]
            }
        ];

        TemplateStore Store()
        {
            var store = new TemplateStore();
            store.RegisterDocxTemplate<TokenFreePartTests.Model>(new MemoryStream(bytes));
            return store;
        }

        var expected = new[]
        {
            await Render(Store(), models[0]),
            await Render(Store(), models[1])
        };

        var store = Store();
        var results = RenderConcurrently(index => Render(store, models[index % models.Length]));

        for (var thread = 0; thread < threads; thread++)
        {
            foreach (var result in results[thread])
            {
                await Assert.That(result.AsSpan().SequenceEqual(expected[thread % models.Length])).IsTrue();
            }
        }
    }

    [Test]
    public async Task ACloneIsTheTreeAParseOfTheSnapshotGives()
    {
        var wordAuthored = new TemplateStore();
        wordAuthored.RegisterDocxTemplate<TokenFreePartTests.Model>(
            new MemoryStream(await File.ReadAllBytesAsync(Path.Combine(ProjectFiles.ProjectDirectory, "Docx", "WordAuthored.docx"))));

        foreach (var snapshot in new[]
                 {
                     ((RegisteredDocxTemplate) ReportStore().Resolve(typeof(Report))).Snapshot,
                     ((RegisteredDocxTemplate) OrderStore().Resolve(typeof(Order))).Snapshot,
                     ((RegisteredDocxTemplate) wordAuthored.Resolve(typeof(TokenFreePartTests.Model))).Snapshot
                 })
        {
            var parsedXml = PartXml(snapshot, _ => { });
            var parts = ParsedParts.Read(snapshot, parsedXml.Keys.ToHashSet());

            // The package a render opens holds those parts with nothing in them.
            foreach (var (_, xml) in PartXml(parts.Bytes, _ => { }))
            {
                await Assert.That(XElement.Parse(xml).HasElements).IsFalse();
            }

            // Filled, changed beyond recognition, and filled again into another package.
            var first = PartXml(
                parts.Bytes,
                doc =>
                {
                    foreach (var (uri, _, root) in DocxCloner.EnumerateParts(doc))
                    {
                        parts.Fill(uri, root);
                        foreach (var text in root.Descendants<Text>().ToList())
                        {
                            text.Text = "changed";
                        }

                        foreach (var paragraph in root.Descendants<Paragraph>().ToList())
                        {
                            paragraph.ParagraphProperties = new(new KeepNext());
                            paragraph.Remove();
                        }
                    }
                });
            var second = PartXml(
                parts.Bytes,
                doc =>
                {
                    foreach (var (uri, _, root) in DocxCloner.EnumerateParts(doc))
                    {
                        parts.Fill(uri, root);
                    }
                });

            await Assert.That(first.Keys).IsEquivalentTo(parsedXml.Keys);
            foreach (var (uri, xml) in parsedXml)
            {
                await Assert.That(first[uri]).IsNotEqualTo(xml);
                await Assert.That(second[uri]).IsEqualTo(xml);
            }
        }
    }

    [Test]
    public async Task ATemplateWithoutTokensKeepsItsSnapshot()
    {
        using var template = DocxTemplateBuilder.Build("No tokens here");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Report>(template);
        var snapshot = ((RegisteredDocxTemplate) store.Resolve(typeof(Report))).Snapshot;

        await Assert.That(ReferenceEquals(ParsedParts.Read(snapshot, new HashSet<string>()).Bytes, snapshot)).IsTrue();

        var first = await Render(store, new Report());
        await Assert.That(await Render(store, new Report())).IsEquivalentTo(first);
        await Assert.That(Text(first, "word/document.xml")).IsEqualTo(Text(snapshot, "word/document.xml"));
    }

    // Each part's root as xml, after change has run over the open package.
    static Dictionary<string, string> PartXml(byte[] package, Action<WordprocessingDocument> change)
    {
        var result = new Dictionary<string, string>();
        using var stream = new MemoryStream();
        stream.Write(package);
        using var doc = WordprocessingDocument.Open(stream, true);
        change(doc);
        foreach (var (uri, _, root) in DocxCloner.EnumerateParts(doc))
        {
            result[uri] = root.OuterXml;
        }

        return result;
    }

    static byte[][][] RenderConcurrently(Func<int, Task<byte[]>> render)
    {
        var results = new byte[threads][][];
        var failures = new List<Exception>();
        using var barrier = new Barrier(threads);
        var workers = new Thread[threads];
        for (var index = 0; index < threads; index++)
        {
            var thread = index;
            workers[thread] = new(() =>
            {
                try
                {
                    var rendered = new byte[rendersPerThread][];
                    barrier.SignalAndWait();
                    for (var count = 0; count < rendersPerThread; count++)
                    {
                        rendered[count] = render(thread).GetAwaiter().GetResult();
                    }

                    results[thread] = rendered;
                }
                catch (Exception exception)
                {
                    lock (failures)
                    {
                        failures.Add(exception);
                    }
                }
            });
            workers[thread].Start();
        }

        foreach (var worker in workers)
        {
            worker.Join();
        }

        if (failures.Count > 0)
        {
            throw new AggregateException(failures);
        }

        return results;
    }

    static async Task<byte[]> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        return stream.ToArray();
    }

    static string Text(byte[] package, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var reader = new StreamReader(archive.GetEntry(name)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    static Order NewOrder(string number, string purchaseOrder, bool approved, string? notes, int budgets) =>
        new()
        {
            Number = number,
            PurchaseOrder = purchaseOrder,
            Approved = approved,
            Notes = notes,
            Budgets = Enumerable.Range(0, budgets)
                .Select(_ => new Budget
                {
                    Year = (2026 + _).ToString(),
                    Amount = 100 * _
                })
                .ToList()
        };

    // One package per template, registered into every store: a package built twice differs in the
    // relationship ids the SDK draws, and so would the renders.
    static readonly byte[] reportTemplate = ReportTemplate();
    static readonly byte[] orderTemplate = OrderTemplate();

    static TemplateStore OrderStore()
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Order>(new MemoryStream(orderTemplate));
        return store;
    }

    static byte[] OrderTemplate()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Order {{ Number }}

            PO: {{ PurchaseOrder }}

            Approved: {{ Approved }}

            {{ Notes }}

            {% for budget in Budgets %}

            {{ budget.Year }}: {{ budget.Amount }}

            {% endfor %}

            End of {{ Number }}
            """);
        return template.ToArray();
    }

    static TemplateStore ReportStore()
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Report>(new MemoryStream(reportTemplate));
        return store;
    }

    // Tokens in the body, a header, a footer, the footnotes and the endnotes: every kind of part a
    // render clones. The body also holds what a clone could plausibly get wrong - a w:smartTag,
    // which the SDK loads as unknown elements, a token a MutateToken edits in place, a table cell,
    // a loop, a conditional and a structural replacement.
    static byte[] ReportTemplate()
    {
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(
                new Body(
                    Paragraph("Title: {{ Title }}"),
                    new Paragraph(
                        $"""<w:p xmlns:w="{word}"><w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>{"{{ Title }}"} in a smart tag</w:t></w:r></w:smartTag></w:p>"""),
                    Paragraph("{% for row in Rows %}"),
                    Paragraph("{{ row.Name }}: {{ row.Amount }}"),
                    Paragraph("{% endfor %}"),
                    Paragraph("{% if Flag %}"),
                    Paragraph("Flagged"),
                    Paragraph("{% else %}"),
                    Paragraph("Not flagged"),
                    Paragraph("{% endif %}"),
                    Paragraph("{{ Notes | markdown }}"),
                    Paragraph("{{ Mark }}"),
                    new Table(
                        new TableProperties(
                            new TableWidth
                            {
                                Type = TableWidthUnitValues.Auto
                            }),
                        new TableGrid(new GridColumn()),
                        new TableRow(new TableCell(Paragraph("Cell {{ Title }}")))),
                    new Paragraph(
                        new Run(new Text("Noted")),
                        new Run(
                            new FootnoteReference
                            {
                                Id = 1
                            }),
                        new Run(
                            new EndnoteReference
                            {
                                Id = 1
                            })),
                    new SectionProperties(
                        new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdHeader"
                        },
                        new FooterReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdFooter"
                        },
                        new PageSize
                        {
                            Width = 6500,
                            Height = 8000
                        })));

            mainPart.AddNewPart<HeaderPart>("rIdHeader").Header = new(
                Paragraph("{{ Title }} header"),
                Paragraph("{% for row in Rows %}"),
                Paragraph("{{ row.Name }}"),
                Paragraph("{% endfor %}"));
            mainPart.AddNewPart<FooterPart>("rIdFooter").Footer = new(
                Paragraph("{% if Flag %}"),
                Paragraph("Flagged footer"),
                Paragraph("{% endif %}"),
                Paragraph("Footer of {{ Title }}"));
            mainPart.AddNewPart<FootnotesPart>("rIdFootnotes").Footnotes = new(
                new Footnote(Paragraph("Footnote on {{ Title }}"))
                {
                    Id = 1
                });
            mainPart.AddNewPart<EndnotesPart>("rIdEndnotes").Endnotes = new(
                new Endnote(Paragraph("Endnote on {{ Title }}"))
                {
                    Id = 1
                });
        }

        return stream.ToArray();
    }

    static Paragraph Paragraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
