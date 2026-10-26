// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// The relationship a link gets: one per target per part, numbered <c>rIdLink{n}</c> after the
/// highest the part holds. The part is not read for every link, so these pin the ids where
/// something other than a markdown link adds a relationship in between: an html block, an image,
/// an Excelsior table, a caller's delegate, a template that arrives with relationships of its own.
/// </summary>
public partial class LinkRelationshipTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public string First { get; init; } = "";
        public string Second { get; init; } = "";
        public string Third { get; init; } = "";
        public string Html { get; init; } = "";

        public TokenValue Token { get; init; } = "";
        public TokenValue Mutation { get; init; } = "";

        [ExcelsiorTable]
        public IReadOnlyList<Row>? Rows { get; init; }

        public IReadOnlyList<string> Items { get; init; } = [];
    }

    public class Row
    {
        [Column(Order = 1)]
        public required Link Site { get; init; }
    }

    const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string onePixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=";

    [Test]
    public async Task MarkdownTemplateSharesARelationshipPerTarget()
    {
        using var stream = await RenderMarkdown(
            """
            [a](https://example.com/a) [b](https://example.com/b) [a again](https://example.com/a)

            <https://example.com/c> and <https://example.com/b>

            - [d](https://example.com/d)
            - [a](https://example.com/a)
            """,
            new());

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/a",
            "rIdLink2 https://example.com/b",
            "rIdLink1 https://example.com/a",
            "rIdLink3 https://example.com/c",
            "rIdLink2 https://example.com/b",
            "rIdLink4 https://example.com/d",
            "rIdLink1 https://example.com/a"
        ]);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(4);
        await AssertSound(doc);
    }

    // The html block's link is added to the main part by OpenXmlHtml, in the middle of the walk.
    // The markdown link after it to the same place shares that relationship, and the image in the
    // last paragraph is another relationship added from outside.
    [Test]
    public async Task MarkdownTemplateSeesWhatAnHtmlBlockAndAnImageAdd()
    {
        using var stream = await RenderMarkdown(
            $"""
            [a](https://example.com/a) [b](https://example.com/b)

            <p>see <a href="https://example.com/html">html</a> and <a href="https://example.com/a">a</a></p>

            [html](https://example.com/html) [c](https://example.com/c) [a](https://example.com/a)

            ![logo](data:image/png;base64,{onePixel}) [d](https://example.com/d) [c](https://example.com/c)
            """,
            new());

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/a",
            "rIdLink2 https://example.com/b",
            "rLink https://example.com/html",
            "rLink1 https://example.com/a",
            "rLink https://example.com/html",
            "rIdLink3 https://example.com/c",
            "rIdLink1 https://example.com/a",
            "rIdLink4 https://example.com/d",
            "rIdLink3 https://example.com/c"
        ]);
        await AssertSound(doc);
    }

    // The tokens are swapped in after the walk, so their relationships join the ones the walk made.
    [Test]
    public async Task MarkdownTemplateWithTokensAndAnExcelsiorTable()
    {
        using var stream = await RenderMarkdown(
            """
            [a](https://example.com/a) [b](https://example.com/b)

            {{ Html | html }}

            {{ Rows }}

            {{ Token }}

            [c](https://example.com/c) [a](https://example.com/a)
            """,
            new()
            {
                Html = """<p><a href="https://example.com/b">b</a> <a href="https://example.com/h">h</a></p>""",
                Rows = [new() { Site = new("https://example.com/e", "Excelsior") }],
                Token = TakesLinkId("rIdLink9", "https://example.com/t")
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/a",
            "rIdLink2 https://example.com/b",
            "rLink https://example.com/b",
            "rLink1 https://example.com/h",
            "rIdLink4 https://example.com/e",
            "rIdLink9 https://example.com/t",
            "rIdLink3 https://example.com/c",
            "rIdLink1 https://example.com/a"
        ]);
        await AssertSound(doc);
    }

    // A delegate takes rIdLink9 between two markdown values. The links after it number on from
    // there, and the one to the delegate's target shares the delegate's relationship.
    [Test]
    public async Task DocxTemplateSeesWhatADelegateAdds()
    {
        using var stream = await RenderDocx(
            Template(body: "{{ First | markdown }}\n{{ Token }}\n{{ Second | markdown }}\nLink: {{ Mutation }}\n{{ Third | markdown }}"),
            new()
            {
                First = "[a](https://example.com/a) [b](https://example.com/b) [c](https://example.com/c)",
                Token = TakesLinkId("rIdLink9", "https://example.com/t"),
                Second = "[t](https://example.com/t) [d](https://example.com/d) [a](https://example.com/a) [e](https://example.com/e)",
                Mutation = new MutateToken((paragraph, context) =>
                {
                    context.MainPart.AddHyperlinkRelationship(new("https://example.com/m"), true, "rIdLink20");
                    paragraph.Append(new Hyperlink(new Run(new Text("mutation"))) { Id = "rIdLink20" });
                }),
                Third = "[m](https://example.com/m) [f](https://example.com/f) [b](https://example.com/b)"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/a",
            "rIdLink2 https://example.com/b",
            "rIdLink3 https://example.com/c",
            "rIdLink9 https://example.com/t",
            "rIdLink9 https://example.com/t",
            "rIdLink10 https://example.com/d",
            "rIdLink1 https://example.com/a",
            "rIdLink11 https://example.com/e",
            "rIdLink20 https://example.com/m",
            "rIdLink20 https://example.com/m",
            "rIdLink21 https://example.com/f",
            "rIdLink2 https://example.com/b"
        ]);
        await AssertSound(doc);
    }

    // The same for what OpenXmlHtml and Excelsior add: a markdown link to a target either of them
    // linked first shares that relationship, whatever id it was given.
    [Test]
    public async Task DocxTemplateSeesWhatHtmlAndExcelsiorAdd()
    {
        using var stream = await RenderDocx(
            Template(body: "{{ First | markdown }}\n{{ Html | html }}\n{{ Second | markdown }}\n{{ Rows }}\n{{ Third | markdown }}"),
            new()
            {
                First = "[a](https://example.com/a) [b](https://example.com/b) [c](https://example.com/c)",
                Html = """<p><a href="https://example.com/h">h</a> <a href="https://example.com/a">a</a></p>""",
                Second = "[h](https://example.com/h) [d](https://example.com/d) [e](https://example.com/e)",
                Rows = [new() { Site = new("https://example.com/x", "Excelsior") }, new() { Site = new("https://example.com/d", "d") }],
                Third = "[x](https://example.com/x) [f](https://example.com/f) [g](https://example.com/g)"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/a",
            "rIdLink2 https://example.com/b",
            "rIdLink3 https://example.com/c",
            "rLink https://example.com/h",
            "rLink1 https://example.com/a",
            "rLink https://example.com/h",
            "rIdLink4 https://example.com/d",
            "rIdLink5 https://example.com/e",
            "rIdLink6 https://example.com/x",
            "rIdLink7 https://example.com/d",
            "rIdLink6 https://example.com/x",
            "rIdLink8 https://example.com/f",
            "rIdLink9 https://example.com/g"
        ]);
        await AssertSound(doc);
    }

    // A template can arrive holding relationships under the ids links are numbered with, and not
    // only hyperlinks: the numbering goes on from the highest of any kind, and a target the
    // template links twice shares the first.
    [Test]
    public async Task TemplateRelationshipsAreNumberedAround()
    {
        using var stream = await RenderDocx(
            Template(
                body: "{{ First | markdown }}\n{{ Second | markdown }}",
                customize: main =>
                {
                    main.AddHyperlinkRelationship(new("https://example.com/own"), true, "rIdLink5");
                    main.AddHyperlinkRelationship(new("https://example.com/own"), true, "rIdLink3");
                    main.AddHyperlinkRelationship(new("https://example.com/other"), true, "rOther");
                    main.AddExternalRelationship(relationships + "/image", new("https://example.com/logo.png"), "rIdLink7");
                    var body = main.Document!.Body!;
                    body.InsertAt(new Paragraph(new Hyperlink(new Run(new Text("own"))) { Id = "rIdLink5" }), 0);
                }),
            new()
            {
                First = "[own](https://example.com/own) [a](https://example.com/a) [other](https://example.com/other)",
                Second = "[b](https://example.com/b) [a](https://example.com/a) [own](https://example.com/own)"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(Links(doc.MainDocumentPart!)).IsEquivalentTo(
        [
            "rIdLink5 https://example.com/own",
            "rIdLink5 https://example.com/own",
            "rIdLink8 https://example.com/a",
            "rOther https://example.com/other",
            "rIdLink9 https://example.com/b",
            "rIdLink8 https://example.com/a",
            "rIdLink5 https://example.com/own"
        ]);
        await AssertSound(doc);
    }

    // Each part numbers its own. The header's html links are added to the main part by OpenXmlHtml
    // and moved across one token at a time; its markdown links are added to it directly.
    [Test]
    public async Task HeaderFooterAndNotesNumberTheirOwn()
    {
        using var stream = await RenderDocx(
            Template(
                body: "{{ First | markdown }}\n{% for item in Items %}\n{{ item | markdown }}\n{% endfor %}",
                header: "{% for item in Items %}\n{{ item | markdown }}\n{% endfor %}\n{{ Html | html }}\n{{ First | markdown }}",
                footer: "{{ Html | html }}\n{{ Second | markdown }}\n{{ Html | html }}\n{{ First | markdown }}",
                footnote: "{{ Rows }}\n{{ Second | markdown }}\nend"),
            new()
            {
                First = "[a](https://example.com/a) [b](https://example.com/b) [h](https://example.com/h)",
                Second = "[x](https://example.com/x) [c](https://example.com/c) [h](https://example.com/h) [d](https://example.com/d)",
                Html = """<p><a href="https://example.com/h">h</a> <a href="https://example.com/i">i</a></p>""",
                Rows = [new() { Site = new("https://example.com/x", "Excelsior") }],
                Items =
                [
                    "[1](https://example.com/1)",
                    "[2](https://example.com/2)",
                    "[1](https://example.com/1) [3](https://example.com/3)"
                ]
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/a",
            "rIdLink2 https://example.com/b",
            "rIdLink3 https://example.com/h",
            "rIdLink4 https://example.com/1",
            "rIdLink5 https://example.com/2",
            "rIdLink4 https://example.com/1",
            "rIdLink6 https://example.com/3"
        ]);
        await Assert.That(Links(main.HeaderParts.Single())).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/1",
            "rIdLink2 https://example.com/2",
            "rIdLink1 https://example.com/1",
            "rIdLink3 https://example.com/3",
            "rIdLink4 https://example.com/h",
            "rIdLink5 https://example.com/i",
            "rIdLink6 https://example.com/a",
            "rIdLink7 https://example.com/b",
            "rIdLink4 https://example.com/h"
        ]);
        await Assert.That(Links(main.FooterParts.Single())).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/h",
            "rIdLink2 https://example.com/i",
            "rIdLink3 https://example.com/x",
            "rIdLink4 https://example.com/c",
            "rIdLink1 https://example.com/h",
            "rIdLink5 https://example.com/d",
            "rIdLink1 https://example.com/h",
            "rIdLink2 https://example.com/i",
            "rIdLink6 https://example.com/a",
            "rIdLink7 https://example.com/b",
            "rIdLink1 https://example.com/h"
        ]);
        await Assert.That(Links(main.FootnotesPart!)).IsEquivalentTo(
        [
            "rIdLink1 https://example.com/x",
            "rIdLink1 https://example.com/x",
            "rIdLink2 https://example.com/c",
            "rIdLink3 https://example.com/h",
            "rIdLink4 https://example.com/d"
        ]);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(6);
        await AssertSound(doc);
    }

    // Enough links that a wrong counter or a missed target shows: every distinct target once, each
    // numbered in the order it was first linked, in a part that already held some.
    [Test]
    public async Task ManyLinksKeepTheirOrder()
    {
        var builder = new StringBuilder();
        var expected = new List<string>();
        for (var index = 0; index < 400; index++)
        {
            var target = index % 7 == 3 ? index / 2 : index;
            builder.AppendLine($"[{index}](https://example.com/{target})");
            builder.AppendLine();
        }

        var ids = new Dictionary<int, int>();
        for (var index = 0; index < 400; index++)
        {
            var target = index % 7 == 3 ? index / 2 : index;
            if (!ids.TryGetValue(target, out var id))
            {
                id = ids.Count + 1;
                ids.Add(target, id);
            }

            expected.Add($"rIdLink{id} https://example.com/{target}");
        }

        using var stream = await RenderMarkdown(builder.ToString(), new());
        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(Links(main)).IsEquivalentTo(expected);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(ids.Count);
        await AssertSound(doc);
    }

    // The index belongs to one render's package. Renders of one registered template at the same
    // time each number their own, and all come out as the bytes a render alone gives.
    [Test]
    public async Task ConcurrentRendersAreByteIdentical()
    {
        using var template = Template(
            body: "{{ First | markdown }}\n{{ Html | html }}\n{% for item in Items %}\n{{ item | markdown }}\n{% endfor %}\n{{ Rows }}\n{{ Second | markdown }}",
            header: "{{ Html | html }}\n{{ First | markdown }}",
            footnote: "{% for item in Items %}\n{{ item | markdown }}\n{% endfor %}");
        var docx = new TemplateStore();
        docx.RegisterDocxTemplate<Doc>(template);

        using var styleSource = Template(header: "{{ First | markdown }}");
        var markdown = new TemplateStore();
        markdown.RegisterMarkdownTemplate<Doc>(
            """
            {{ First | markdown }}

            {% for item in Items %}{{ item | markdown }} {% endfor %}

            <p><a href="https://example.com/h">h</a></p>

            {{ Rows }}

            [a](https://example.com/a) [z](https://example.com/z)
            """,
            styleSource);

        var model = new Doc
        {
            First = "[a](https://example.com/a) [b](https://example.com/b) [h](https://example.com/h)",
            Second = "[x](https://example.com/x) [c](https://example.com/c) [a](https://example.com/a)",
            Html = """<p><a href="https://example.com/h">h</a> <a href="https://example.com/i">i</a></p>""",
            Rows = [new() { Site = new("https://example.com/x", "Excelsior") }],
            Items = Enumerable.Range(0, 60).Select(_ => $"[{_}](https://example.com/{_ % 40})").ToList()
        };

        foreach (var store in new[] { docx, markdown })
        {
            var alone = await Bytes(store, model);
            var renders = await Task.WhenAll(
                Enumerable.Range(0, 24).Select(_ => Task.Run(() => Bytes(store, model))));
            foreach (var render in renders)
            {
                await Assert.That(render.AsSpan().SequenceEqual(alone)).IsTrue();
            }

            using var doc = WordprocessingDocument.Open(new MemoryStream(alone), false);
            await AssertSound(doc);
        }
    }

    static OpenXmlToken TakesLinkId(string id, string target) =>
        new(context =>
        {
            context.MainPart.AddHyperlinkRelationship(new(target), true, id);
            return [new Paragraph(new Hyperlink(new Run(new Text("token"))) { Id = id })];
        });

    static async Task<byte[]> Bytes(TemplateStore store, Doc model)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        return stream.ToArray();
    }

    static async Task<MemoryStream> RenderMarkdown(string markdown, Doc model)
    {
        using var styleSource = Template();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return new(await Bytes(store, model));
    }

    static async Task<MemoryStream> RenderDocx(MemoryStream template, Doc model)
    {
        using (template)
        {
            var store = new TemplateStore();
            store.RegisterDocxTemplate<Doc>(template);
            return new(await Bytes(store, model));
        }
    }

    // Each hyperlink of the part in document order, as its relationship id and where that leads.
    static List<string> Links(OpenXmlPart part) =>
        part.RootElement!
            .Descendants<Hyperlink>()
            .Where(_ => _.Id != null)
            .Select(link => $"{link.Id!.Value} {part.HyperlinkRelationships.Single(_ => _.Id == link.Id!.Value).Uri.OriginalString}")
            .ToList();

    static List<string> Targets(List<string> links) =>
        links
            .Select(_ => _[(_.IndexOf(' ') + 1)..])
            .ToList();

    // Every relationship attribute resolves in the part that holds it, no part holds two
    // relationships under one id, and the document is valid besides.
    static async Task AssertSound(WordprocessingDocument doc)
    {
        var problems = new List<string>();
        var main = doc.MainDocumentPart!;
        List<OpenXmlPart> parts =
        [
            main,
            ..main.HeaderParts,
            ..main.FooterParts
        ];
        if (main.FootnotesPart is { } footnotes)
        {
            parts.Add(footnotes);
        }

        foreach (var part in parts)
        {
            var all = part.Parts.Select(_ => _.RelationshipId)
                .Concat(part.ExternalRelationships.Select(_ => _.Id))
                .Concat(part.HyperlinkRelationships.Select(_ => _.Id))
                .ToList();
            var ids = new HashSet<string>(all);
            if (ids.Count != all.Count)
            {
                problems.Add($"{part.Uri}: a relationship id is held twice");
            }

            foreach (var element in part.RootElement!.Descendants())
            {
                foreach (var attribute in element.GetAttributes())
                {
                    if (attribute.NamespaceUri == relationships &&
                        !ids.Contains(attribute.Value!))
                    {
                        problems.Add($"{part.Uri}: {element.LocalName} {attribute.LocalName}=\"{attribute.Value}\"");
                    }
                }
            }
        }

        problems.AddRange(
            new OpenXmlValidator(FileFormatVersions.Microsoft365)
                .Validate(doc)
                .Select(_ => $"{_.Part?.Uri} {_.Path?.XPath}: {_.Description}"));
        await Assert.That(problems).IsEmpty();
    }

    // A paragraph per line of text in each part asked for.
    static MemoryStream Template(
        string body = "Body",
        string? header = null,
        string? footer = null,
        string? footnote = null,
        Action<MainDocumentPart>? customize = null)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            // Excelsior adds a styles part, under a random id, to a document that has none.
            mainPart.AddNewPart<StyleDefinitionsPart>("rIdStyles").Styles = new();

            var section = new SectionProperties();
            if (header != null)
            {
                mainPart.AddNewPart<HeaderPart>("rIdHeader").Header = new(Paragraphs(header));
                section.Append(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    });
            }

            if (footer != null)
            {
                mainPart.AddNewPart<FooterPart>("rIdFooter").Footer = new(Paragraphs(footer));
                section.Append(
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdFooter"
                    });
            }

            var notes = new Paragraph(new Run(new Text("Notes")));
            if (footnote != null)
            {
                mainPart.AddNewPart<FootnotesPart>("rIdFootnotes").Footnotes = new(
                    new Footnote(Paragraphs(footnote))
                    {
                        Id = 1
                    });
                notes.Append(
                    new Run(
                        new FootnoteReference
                        {
                            Id = 1
                        }));
            }

            section.Append(
                new PageSize
                {
                    Width = 6500,
                    Height = 8000
                });

            var documentBody = new Body(Paragraphs(body));
            documentBody.Append(notes, section);
            mainPart.Document = new(documentBody);
            customize?.Invoke(mainPart);
        }

        stream.Position = 0;
        return stream;
    }

    static IEnumerable<OpenXmlElement> Paragraphs(string text) =>
        text
            .Split('\n')
            .Select(_ => new Paragraph(new Run(new Text(_) { Space = SpaceProcessingModeValues.Preserve })));
}
