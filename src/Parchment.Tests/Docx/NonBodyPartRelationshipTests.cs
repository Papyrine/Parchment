using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A relationship id resolves against the part holding the element that carries it. A link or an
/// image a token produces in a header, footer, footnote or endnote therefore has to hang off that
/// part: left on the main document part the id resolves nowhere, and Word reports the file as
/// corrupt.
/// </summary>
public partial class NonBodyPartRelationshipTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Doc
    {
        public string Title { get; init; } = "Report";

        [Html]
        public string Html { get; init; } = "";

        [Markdown]
        public string Markdown { get; init; } = "";

        // Unannotated, for the `| html` and `| markdown` filters.
        public string PlainHtml { get; init; } = "";
        public string PlainMarkdown { get; init; } = "";

        public TokenValue Token { get; init; } = "";
        public TokenValue Mutation { get; init; } = "";

        [ExcelsiorTable]
        public IReadOnlyList<Row>? Rows { get; init; }

        public IReadOnlyList<string> Sites { get; init; } = [];
    }

    public class Row
    {
        [Column(Order = 1)]
        public required Link Site { get; init; }
    }

    const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string imageRelationship = relationships + "/image";

    // 1x1 and 2x2, so two images in one part are two image parts.
    const string onePixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=";
    const string twoPixels = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP4z8AARAwQCgAf7gP9i18U1AAAAABJRU5ErkJggg==";

    [Test]
    public async Task HtmlLinkBelongsToThePartHoldingTheToken()
    {
        using var stream = await Render(
            Template(
                body: "{{ Html }}",
                header: "{{ Html }}",
                footer: "{{ Html }}",
                footnote: "{{ Html }}",
                endnote: "{{ Html }}"),
            new()
            {
                Html = """<p>see <a href="https://example.com/a">the site</a></p>"""
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        foreach (var part in ContentParts(doc))
        {
            await Assert.That(LinkTargets(part)).IsEquivalentTo(["https://example.com/a"]);
            await Assert.That(part.HyperlinkRelationships.Count()).IsEqualTo(1);
        }

        await AssertSound(doc);
    }

    // Alone in the paragraph, which replaces it, and sharing it, which splices into it.
    [Test]
    public async Task FilterLinkBelongsToThePartHoldingTheToken()
    {
        using var stream = await Render(
            Template(
                header: "{{ PlainHtml | html }}\nSee: {{ PlainHtml | html }}",
                footnote: "{{ PlainMarkdown | markdown }}\nSee: {{ PlainMarkdown | markdown }}"),
            new()
            {
                PlainHtml = """<a href="https://example.com/h">html</a>""",
                PlainMarkdown = "[markdown](https://example.com/m)"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;

        await Assert.That(LinkTargets(main.HeaderParts.Single()))
            .IsEquivalentTo(["https://example.com/h", "https://example.com/h"]);
        await Assert.That(LinkTargets(main.FootnotesPart!))
            .IsEquivalentTo(["https://example.com/m", "https://example.com/m"]);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(0);
        await AssertSound(doc);
    }

    // A markdown link shares one relationship per target. The body and the header link to the same
    // place here, and each has to end up with a relationship of its own.
    [Test]
    public async Task MarkdownLinkBelongsToThePartHoldingTheToken()
    {
        using var stream = await Render(
            Template(
                body: "{{ Markdown }}",
                header: "{{ Markdown }}\nAgain: {{ PlainMarkdown | markdown }}",
                endnote: "{{ Markdown }}"),
            new()
            {
                Markdown = "See [site](https://example.com/a) or <https://example.com/b>",
                PlainMarkdown = "[site](https://example.com/a)"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var header = main.HeaderParts.Single();

        await Assert.That(LinkTargets(main)).IsEquivalentTo(["https://example.com/a", "https://example.com/b"]);
        await Assert.That(LinkTargets(main.EndnotesPart!)).IsEquivalentTo(["https://example.com/a", "https://example.com/b"]);
        await Assert.That(LinkTargets(header))
            .IsEquivalentTo(["https://example.com/a", "https://example.com/b", "https://example.com/a"]);
        await Assert.That(header.HyperlinkRelationships.Count()).IsEqualTo(2);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(2);
        await AssertSound(doc);
    }

    // Two html images in the one header: OpenXmlHtml numbers each drawing from the images the main
    // part holds, so the second still has to come out with an id of its own.
    [Test]
    public async Task ImageBelongsToThePartHoldingTheToken()
    {
        using var stream = await Render(
            Template(
                body: "{{ Markdown }}",
                header: "{{ Html }}\n{{ PlainHtml | html }}",
                footer: "{{ Markdown }}"),
            new()
            {
                Html = $"""<p><img src="data:image/png;base64,{onePixel}" width="10" height="10"/></p>""",
                PlainHtml = $"""<p><img src="data:image/png;base64,{twoPixels}" width="10" height="10"/></p>""",
                Markdown = $"![logo](data:image/png;base64,{onePixel})"
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;

        await Assert.That(Images(main)).IsEquivalentTo([onePixel]);
        await Assert.That(Images(main.HeaderParts.Single())).IsEquivalentTo([onePixel, twoPixels]);
        await Assert.That(Images(main.FooterParts.Single())).IsEquivalentTo([onePixel]);
        await Assert.That(main.ImageParts.Count()).IsEqualTo(1);
        await AssertSound(doc);
    }

    // An svg is two image parts, the svg and a raster fallback, and the svg is referred to from an
    // extension element the SDK has no type for.
    [Test]
    public async Task SvgImageBelongsToThePartHoldingTheToken()
    {
        var svg = Convert.ToBase64String("""<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>"""u8);
        using var stream = await Render(
            Template(footnote: "{{ Html }}"),
            new()
            {
                Html = $"""<p><img src="data:image/svg+xml;base64,{svg}" width="10" height="10"/></p>"""
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;

        await Assert.That(main.FootnotesPart!.ImageParts.Select(_ => _.ContentType))
            .IsEquivalentTo(["image/png", "image/svg+xml"]);
        await Assert.That(main.ImageParts.Count()).IsEqualTo(0);
        await AssertSound(doc);
    }

    [Test]
    public async Task ContextImagePartBelongsToThePartHoldingTheToken()
    {
        var bytes = Convert.FromBase64String(onePixel);
        using var stream = await Render(
            Template(
                body: "{{ Token }}",
                header: "{{ Token }}",
                footer: "Logo: {{ Mutation }}",
                footnote: "{{ Token }}"),
            new()
            {
                Token = new OpenXmlToken(context => [new Paragraph(new Run(Picture(context.AddImagePart(bytes, "image/png"))))]),
                Mutation = new MutateToken((paragraph, context) => paragraph.Append(new Run(Picture(context.AddImagePart(bytes, "image/png")))))
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        foreach (var part in ContentParts(doc))
        {
            await Assert.That(Images(part)).IsEquivalentTo([onePixel]);
        }

        await Assert.That(doc.MainDocumentPart!.ImageParts.Count()).IsEqualTo(1);
        await AssertSound(doc);
    }

    // The context hands a delegate the main part and nothing else, so a delegate that builds its
    // own link, linked picture or embedded part can only add the relationship there.
    [Test]
    public async Task RelationshipADelegateAddsToTheMainPartIsMoved()
    {
        using var stream = await Render(
            Template(
                header: "{{ Token }}",
                footer: "Link: {{ Mutation }}"),
            new()
            {
                Token = new OpenXmlToken(context =>
                {
                    var mainPart = context.MainPart;
                    var link = mainPart.AddHyperlinkRelationship(new("https://example.com/t"), true).Id;
                    var picture = mainPart.AddExternalRelationship(imageRelationship, new("https://example.com/logo.png")).Id;
                    var chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
                    using (var html = new MemoryStream("<html><body><p>chunk</p></body></html>"u8.ToArray()))
                    {
                        chunk.FeedData(html);
                    }

                    return
                    [
                        new Paragraph(new Hyperlink(new Run(new Text("token"))) { Id = link }),
                        new Paragraph(new Run(Picture(picture, linked: true))),
                        new AltChunk { Id = mainPart.GetIdOfPart(chunk) }
                    ];
                }),
                Mutation = new MutateToken((paragraph, context) =>
                {
                    var id = context.MainPart.AddHyperlinkRelationship(new("https://example.com/m"), true).Id;
                    paragraph.Append(new Hyperlink(new Run(new Text("mutation"))) { Id = id });
                })
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var header = main.HeaderParts.Single();

        await Assert.That(LinkTargets(header)).IsEquivalentTo(["https://example.com/t"]);
        await Assert.That(header.ExternalRelationships.Select(_ => _.Uri.OriginalString)).IsEquivalentTo(["https://example.com/logo.png"]);
        await Assert.That(header.AlternativeFormatImportParts.Count()).IsEqualTo(1);
        await Assert.That(LinkTargets(main.FooterParts.Single())).IsEquivalentTo(["https://example.com/m"]);

        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(0);
        await Assert.That(main.ExternalRelationships.Count()).IsEqualTo(0);
        await Assert.That(main.AlternativeFormatImportParts.Count()).IsEqualTo(0);
        await AssertSound(doc);
    }

    // A paragraph can hold a link of the template's beside the token. What a MutateToken finds in
    // its host is left as it is, whatever id it carries.
    [Test]
    public async Task MutateTokenLeavesWhatItsHostAlreadyReferredTo()
    {
        using var stream = await Render(
            Template(
                header: "{{ Html }}\nLink: {{ Mutation }}",
                customize: main =>
                {
                    var header = main.HeaderParts.Single();
                    header.AddHyperlinkRelationship(new("https://example.com/own"), true, "rLink");
                    header.Header!.Elements<Paragraph>().Last().Append(new Hyperlink(new Run(new Text("own"))) { Id = "rLink" });
                }),
            new()
            {
                // OpenXmlHtml adds this to the main part as rLink, the id the template's link has.
                Html = """<p><a href="https://example.com/a">the site</a></p>""",
                Mutation = new MutateToken((paragraph, context) =>
                {
                    var id = context.MainPart.AddHyperlinkRelationship(new("https://example.com/m"), true).Id;
                    paragraph.Append(new Hyperlink(new Run(new Text("mutation"))) { Id = id });
                })
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(LinkTargets(doc.MainDocumentPart!.HeaderParts.Single()))
            .IsEquivalentTo(["https://example.com/a", "https://example.com/own", "https://example.com/m"]);
        await AssertSound(doc);
    }

    // Through [ExcelsiorTable], and through the readme's route for a table the attribute cannot
    // express: an OpenXmlToken handing IOpenXmlContext.MainPart to WordTableBuilder.
    [Test]
    public async Task ExcelsiorLinkCellBelongsToThePartHoldingTheToken()
    {
        Row[] rows = [new() { Site = new("https://example.com/e", "Excelsior") }];
        using var stream = await Render(
            Template(
                body: "{{ Rows }}",
                header: "{{ Rows }}\nend",
                footer: "{{ Token }}\nend"),
            new()
            {
                Rows = rows,
                Token = new OpenXmlToken(context => [new WordTableBuilder<Row>(rows).Build(context.MainPart)])
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        foreach (var part in ContentParts(doc))
        {
            await Assert.That(LinkTargets(part)).IsEquivalentTo(["https://example.com/e"]);
            await Assert.That(part.HyperlinkRelationships.Count()).IsEqualTo(1);
        }

        await AssertSound(doc);
    }

    [Test]
    public async Task LoopGivesEveryIterationItsRelationship()
    {
        using var stream = await Render(
            Template(header: "{% for site in Sites %}\n{{ site | html }}\n{% endfor %}"),
            new()
            {
                Sites =
                [
                    """<a href="https://example.com/a">a</a>""",
                    """<a href="https://example.com/b">b</a>""",
                    """<a href="https://example.com/a">a again</a>"""
                ]
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var header = main.HeaderParts.Single();

        await Assert.That(LinkTargets(header))
            .IsEquivalentTo(["https://example.com/a", "https://example.com/b", "https://example.com/a"]);
        await Assert.That(header.HyperlinkRelationships.Count()).IsEqualTo(2);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(0);
        await AssertSound(doc);
    }

    // OpenXmlHtml names its first link rLink on the main part. A header that arrives owning an
    // rLink of its own - a template OpenXmlHtml wrote, say - keeps it, and keeps its target.
    [Test]
    public async Task RelationshipTheTemplateHoldsUnderTheSameIdIsLeftAlone()
    {
        using var stream = await Render(
            Template(
                header: "{{ Html }}",
                customize: main =>
                {
                    var header = main.HeaderParts.Single();
                    header.AddHyperlinkRelationship(new("https://example.com/own"), true, "rLink");
                    header.Header!.Append(new Paragraph(new Hyperlink(new Run(new Text("own"))) { Id = "rLink" }));
                }),
            new()
            {
                Html = """<p><a href="https://example.com/a">the site</a></p>"""
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(LinkTargets(doc.MainDocumentPart!.HeaderParts.Single()))
            .IsEquivalentTo(["https://example.com/a", "https://example.com/own"]);
        await AssertSound(doc);
    }

    // The markdown flow binds a style source's header and footer through the same runner.
    [Test]
    public async Task MarkdownFlowLinkBelongsToThePartHoldingTheToken()
    {
        using var styleSource = Template(
            header: "{{ PlainMarkdown | markdown }}",
            footer: "{{ PlainHtml | html }}");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("# {{ Title }}\n\nSee [site](https://example.com/a)", styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new Doc
            {
                PlainMarkdown = "Header [site](https://example.com/a)",
                PlainHtml = """<a href="https://example.com/f">footer</a>"""
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;

        await Assert.That(LinkTargets(main)).IsEquivalentTo(["https://example.com/a"]);
        await Assert.That(LinkTargets(main.HeaderParts.Single())).IsEquivalentTo(["https://example.com/a"]);
        await Assert.That(LinkTargets(main.FooterParts.Single())).IsEquivalentTo(["https://example.com/f"]);
        await Assert.That(main.HyperlinkRelationships.Count()).IsEqualTo(1);
        await AssertSound(doc);
    }

    // Excelsior adds its link under a random id. Moved to the footer it takes one of Parchment's,
    // so the random one does not reach the output.
    [Test]
    public async Task RenderIsByteIdentical()
    {
        using var template = Template(
            header: "{{ Html }}\n{{ Markdown }}",
            footer: "{{ Rows }}\nend",
            footnote: "{{ PlainHtml | html }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        var model = new Doc
        {
            Html = $"""<p><a href="https://example.com/a">a</a> <img src="data:image/png;base64,{onePixel}" width="10" height="10"/></p>""",
            Markdown = $"[b](https://example.com/b) ![logo](data:image/png;base64,{twoPixels})",
            PlainHtml = """<a href="https://example.com/n">note</a>""",
            Rows = [new() { Site = new("https://example.com/e", "Excelsior") }]
        };

        using var first = new MemoryStream();
        await store.Render(model, first);
        using var second = new MemoryStream();
        await store.Render(model, second);

        await Assert.That(first.ToArray()).IsEquivalentTo(second.ToArray());

        first.Position = 0;
        using var doc = WordprocessingDocument.Open(first, false);
        await AssertSound(doc);
    }

    static async Task<MemoryStream> Render(MemoryStream template, Doc model)
    {
        using (template)
        {
            var store = new TemplateStore();
            store.RegisterDocxTemplate<Doc>(template);

            var stream = new MemoryStream();
            await store.Render(model, stream);
            stream.Position = 0;
            return stream;
        }
    }

    // Every r:id, r:embed and r:link has to resolve in the part that holds it, and the document
    // has to be valid besides.
    static async Task AssertSound(WordprocessingDocument doc)
    {
        var dangling = new List<string>();
        foreach (var part in ContentParts(doc))
        {
            var ids = new HashSet<string>(part.Parts.Select(_ => _.RelationshipId));
            ids.UnionWith(part.ExternalRelationships.Select(_ => _.Id));
            ids.UnionWith(part.HyperlinkRelationships.Select(_ => _.Id));

            foreach (var element in part.RootElement!.Descendants())
            {
                foreach (var attribute in element.GetAttributes())
                {
                    if (attribute.NamespaceUri == relationships &&
                        !ids.Contains(attribute.Value!))
                    {
                        dangling.Add($"{part.Uri}: {element.LocalName} {attribute.LocalName}=\"{attribute.Value}\"");
                    }
                }
            }
        }

        await Assert.That(dangling).IsEmpty();

        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Part?.Uri} {_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static IEnumerable<OpenXmlPart> ContentParts(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart!;
        yield return main;

        foreach (var header in main.HeaderParts)
        {
            yield return header;
        }

        foreach (var footer in main.FooterParts)
        {
            yield return footer;
        }

        if (main.FootnotesPart is { } footnotes)
        {
            yield return footnotes;
        }

        if (main.EndnotesPart is { } endnotes)
        {
            yield return endnotes;
        }
    }

    // Where each hyperlink in the part leads, resolved through the part's own relationships.
    static List<string> LinkTargets(OpenXmlPart part) =>
        part.RootElement!
            .Descendants<Hyperlink>()
            .Select(link => part.HyperlinkRelationships.Single(_ => _.Id == link.Id!.Value).Uri.OriginalString)
            .ToList();

    // The bytes behind each picture in the part, resolved the same way.
    static List<string> Images(OpenXmlPart part) =>
        part.RootElement!
            .Descendants<A.Blip>()
            .Select(blip =>
            {
                using var source = part.GetPartById(blip.Embed!.Value!).GetStream();
                using var buffer = new MemoryStream();
                source.CopyTo(buffer);
                return Convert.ToBase64String(buffer.ToArray());
            })
            .ToList();

    // A paragraph per line of text in each part asked for, and a reference in the body for each
    // kind of note.
    static MemoryStream Template(
        string body = "Body",
        string? header = null,
        string? footer = null,
        string? footnote = null,
        string? endnote = null,
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

            if (endnote != null)
            {
                mainPart.AddNewPart<EndnotesPart>("rIdEndnotes").Endnotes = new(
                    new Endnote(Paragraphs(endnote))
                    {
                        Id = 1
                    });
                notes.Append(
                    new Run(
                        new EndnoteReference
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

    static A.Blip Blip(string relationshipId, bool linked)
    {
        if (linked)
        {
            return new()
            {
                Link = relationshipId
            };
        }

        return new()
        {
            Embed = relationshipId
        };
    }

    static Drawing Picture(string relationshipId, bool linked = false) =>
        new(
            new DW.Inline(
                new DW.Extent
                {
                    Cx = 100000L,
                    Cy = 100000L
                },
                new DW.DocProperties
                {
                    Id = 1U,
                    Name = "Logo"
                },
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties
                                {
                                    Id = 0U,
                                    Name = "logo.png"
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                Blip(relationshipId, linked),
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset
                                    {
                                        X = 0L,
                                        Y = 0L
                                    },
                                    new A.Extents
                                    {
                                        Cx = 100000L,
                                        Cy = 100000L
                                    }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });
}
