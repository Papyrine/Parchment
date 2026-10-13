using DocumentFormat.OpenXml.Validation;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A <c>[SECTION]</c> marker inside a markdown value the docx flow renders. The markdown walk only
/// records what the marker asked for, so the break has to be finished from the section of the
/// template the value landed in - the way the markdown flow finishes one from its style source.
/// </summary>
public partial class MarkdownValueSectionBreakTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Model
    {
        [Markdown]
        public string Body { get; init; } = "";

        public string Text { get; init; } = "";
        public MarkdownToken Token { get; init; } = new("");
        public OpenXmlToken Raw { get; init; } = OpenXmlToken.Empty;
        public List<string> Items { get; init; } = [];
        public bool Show { get; init; }
    }

    const string landscape = "[SECTION]{orientation=landscape}";
    const string aroundLandscape = $"before\n\n{landscape}\n\nwide";

    // The section a marker interrupts keeps everything the template gave it, and the one the marker
    // opens is that section turned on its side. A w:sectPr holding only the orientation - what the
    // markdown walk writes - is a section with no page size, no margins and no header.
    [Test]
    public async Task MarkerIsFinishedFromTheTemplatesSection()
    {
        using var template = Template("Intro\n\n{{ Body }}\n\nOutro");
        using var doc = await Render(
            template,
            new()
            {
                Body = aroundLandscape
            });

        var sections = Sections(doc);
        await Assert.That(sections.Count).IsEqualTo(2);

        var first = sections[0].GetFirstChild<PageSize>()!;
        await Assert.That(first.Width!.Value).IsEqualTo(6500u);
        await Assert.That(first.Height!.Value).IsEqualTo(8000u);
        await Assert.That(first.Orient).IsNull();

        var second = sections[1].GetFirstChild<PageSize>()!;
        await Assert.That(second.Width!.Value).IsEqualTo(8000u);
        await Assert.That(second.Height!.Value).IsEqualTo(6500u);
        await Assert.That(second.Orient!.Value).IsEqualTo(PageOrientationValues.Landscape);

        foreach (var section in sections)
        {
            await Assert.That(section.GetFirstChild<HeaderReference>()!.Id!.Value).IsEqualTo("rIdHeader");
            await Assert.That(section.GetFirstChild<FooterReference>()!.Id!.Value).IsEqualTo("rIdFooter");
            await Assert.That(section.GetFirstChild<PageMargin>()!.Header!.Value).IsEqualTo(720u);
            await Assert.That(section.GetFirstChild<PageBorders>()).IsNotNull();
        }

        // The break sits where the marker was, between the two halves of the value.
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(string.Join('|', body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("Intro|before||wide|Outro");
        await Assert.That(body.Elements<Paragraph>().ElementAt(2).ParagraphProperties!.SectionProperties).IsNotNull();

        await AssertValid(doc);
    }

    // The annotation, the filter and the token type are three ways to say the same thing.
    [Test]
    [Arguments("{{ Body }}")]
    [Arguments("{{ Text | markdown }}")]
    [Arguments("{{ Token }}")]
    public async Task EveryRouteToAMarkdownValueIsFinished(string token)
    {
        using var template = Template($"Intro\n\n{token}\n\nOutro");
        using var doc = await Render(
            template,
            new()
            {
                Body = aroundLandscape,
                Text = aroundLandscape,
                Token = new(aroundLandscape)
            });

        var sections = Sections(doc);
        await Assert.That(sections.Count).IsEqualTo(2);
        await Assert.That(sections[0].GetFirstChild<PageSize>()!.Width!.Value).IsEqualTo(6500u);
        await Assert.That(sections[1].GetFirstChild<PageSize>()!.Width!.Value).IsEqualTo(8000u);
    }

    [Test]
    public async Task OnlyTheFirstSectionRestartsPageNumbering()
    {
        using var template = Template(
            "Intro\n\n{{ Body }}\n\nOutro",
            mainPart => mainPart.Document!.Body!
                .Elements<SectionProperties>()
                .Single()
                .AppendChild(
                    new PageNumberType
                    {
                        Start = 3
                    }));
        using var doc = await Render(
            template,
            new()
            {
                Body = aroundLandscape
            });

        var sections = Sections(doc);
        await Assert.That(sections[0].GetFirstChild<PageNumberType>()!.Start!.Value).IsEqualTo(3);
        await Assert.That(sections[1].GetFirstChild<PageNumberType>()).IsNull();
        await AssertValid(doc);
    }

    // "From here on" stops at the section break the template already had: the section after it
    // keeps the page setup the template gave it.
    [Test]
    public async Task MarkerReachesNoFurtherThanTheTemplatesNextSectionBreak()
    {
        using var template = TwoSectionTemplate();
        using var doc = await Render(
            template,
            new()
            {
                Body = aroundLandscape
            });

        var sections = Sections(doc);
        await Assert.That(DescribeAll(sections)).IsEqualTo("5000x7000 | 7000x5000 landscape | 6500x8000");
        await AssertValid(doc);
    }

    // And a marker in a later section derives from that section, not from the first or the last.
    [Test]
    public async Task MarkerDerivesFromTheSectionItLandsIn()
    {
        using var template = TwoSectionTemplate();
        using var doc = await Render(
            template,
            new()
            {
                Token = new(aroundLandscape)
            });

        var sections = Sections(doc);
        await Assert.That(DescribeAll(sections)).IsEqualTo("5000x7000 | 6500x8000 | 8000x6500 landscape");
        await Assert.That(sections[0].GetFirstChild<HeaderReference>()).IsNull();
        await Assert.That(sections[1].GetFirstChild<HeaderReference>()).IsNotNull();
        await Assert.That(sections[2].GetFirstChild<HeaderReference>()).IsNotNull();
        await AssertValid(doc);
    }

    [Test]
    public async Task MarkersInBothSectionsEachDeriveFromTheirOwn()
    {
        using var template = TwoSectionTemplate();
        using var doc = await Render(
            template,
            new()
            {
                Body = aroundLandscape,
                Token = new(aroundLandscape)
            });

        var sections = Sections(doc);
        await Assert.That(DescribeAll(sections)).IsEqualTo("5000x7000 | 7000x5000 landscape | 6500x8000 | 8000x6500 landscape");
        await AssertValid(doc);
    }

    // A loop body is rendered into a scratch container and moved, once per iteration, so the breaks
    // can only be finished once every iteration is in place.
    [Test]
    public async Task MarkerInsideALoopIsASectionPerIteration()
    {
        using var template = Template("{% for item in Items %}\n\n{{ item | markdown }}\n\n{% endfor %}\n\nOutro");
        using var doc = await Render(
            template,
            new()
            {
                Items = [aroundLandscape, $"tall\n\n[SECTION]{{orientation=portrait}}\n\nafter"]
            });

        var sections = Sections(doc);
        await Assert.That(DescribeAll(sections)).IsEqualTo("6500x8000 | 8000x6500 landscape | 6500x8000");
        await AssertValid(doc);
    }

    [Test]
    public async Task MarkerInsideAConditionalIsFinished()
    {
        using var template = Template("{% if Show %}\n\n{{ Body }}\n\n{% endif %}\n\nOutro");
        using var doc = await Render(
            template,
            new()
            {
                Show = true,
                Body = aroundLandscape
            });

        await Assert.That(DescribeAll(Sections(doc))).IsEqualTo("6500x8000 | 8000x6500 landscape");
        await AssertValid(doc);
    }

    // A value that is nothing but the marker is one paragraph, which a token sharing its paragraph
    // would otherwise unwrap into the host - dropping the paragraph properties that are the break.
    [Test]
    public async Task MarkerSharingItsParagraphSplitsTheHost()
    {
        using var template = Template("Lead {{ Body }} tail");
        using var doc = await Render(
            template,
            new()
            {
                Body = landscape
            });

        var sections = Sections(doc);
        await Assert.That(DescribeAll(sections)).IsEqualTo("6500x8000 | 8000x6500 landscape");

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(string.Join('|', body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("Lead || tail");
        await AssertValid(doc);
    }

    [Test]
    public async Task MarkerAmongBlocksSharingAParagraphIsFinished()
    {
        using var template = Template("Lead {{ Body }} tail");
        using var doc = await Render(
            template,
            new()
            {
                Body = aroundLandscape
            });

        await Assert.That(DescribeAll(Sections(doc))).IsEqualTo("6500x8000 | 8000x6500 landscape");
        await AssertValid(doc);
    }

    // A section break the template or the caller's own code wrote is whole already.
    [Test]
    public async Task SectionBreakFromAnOpenXmlTokenIsLeftAsWritten()
    {
        using var template = Template("Intro\n\n{{ Raw }}\n\nOutro");
        using var doc = await Render(
            template,
            new()
            {
                Raw = new(
                    _ =>
                    [
                        new Paragraph(
                            new ParagraphProperties(
                                new SectionProperties(
                                    new PageSize
                                    {
                                        Width = 4000,
                                        Height = 3000,
                                        Orient = PageOrientationValues.Landscape
                                    })))
                    ])
            });

        await Assert.That(DescribeAll(Sections(doc))).IsEqualTo("4000x3000 landscape | 6500x8000");
    }

    // Word has no section that starts inside a table cell, so there is nothing to finish the break
    // into - and a w:sectPr left in the cell is not a document Word opens.
    [Test]
    public async Task MarkerInATableCellIsRejected()
    {
        using var template = DocxTemplateBuilder.BuildWithTable("Label", "{{ Body }}");

        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () =>
            {
                using var doc = await Render(
                    template,
                    new()
                    {
                        Body = aroundLandscape
                    });
            });

        await Assert.That(exception!.Message).Contains("[SECTION]");
        await Assert.That(exception.Message).Contains("top level");
    }

    // The same goes for a cell of a table the value built itself, which is as far as the markdown
    // flow lets a marker in a grid table get too.
    [Test]
    public async Task MarkerInATableTheValueBuiltIsRejected()
    {
        using var template = Template("Intro\n\n{{ Body }}\n\nOutro");

        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () =>
            {
                using var doc = await Render(
                    template,
                    new()
                    {
                        Body = $"+-----+\n| {landscape} |\n+-----+\n"
                    });
            });

        await Assert.That(exception!.Message).Contains("top level");
    }

    [Test]
    public async Task MarkerInAHeaderIsRejected()
    {
        using var template = Template("Intro", headerText: "{{ Token }}");

        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () =>
            {
                using var doc = await Render(
                    template,
                    new()
                    {
                        Token = new(aroundLandscape)
                    });
            });

        await Assert.That(exception!.Message).Contains("[SECTION]");
        await Assert.That(exception.PartUri).IsEqualTo("/word/header1.xml");
    }

    [Test]
    public async Task MarkerInAFootnoteIsRejected()
    {
        using var template = Template(
            "Intro",
            mainPart =>
            {
                var footnotesPart = mainPart.AddNewPart<FootnotesPart>("rIdFootnotes");
                footnotesPart.Footnotes = new(
                    new Footnote(new Paragraph(new Run(new Text("{{ Token }}"))))
                    {
                        Id = 1
                    });
            });

        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () =>
            {
                using var doc = await Render(
                    template,
                    new()
                    {
                        Token = new(aroundLandscape)
                    });
            });

        await Assert.That(exception!.Message).Contains("[SECTION]");
        await Assert.That(exception.PartUri).IsEqualTo("/word/footnotes.xml");
    }

    // The markdown flow binds a style source's header through the same runner, so the same value
    // reaches the same place there.
    [Test]
    public async Task MarkerInAMarkdownFlowHeaderIsRejected()
    {
        using var styleSource = Template("replaced", headerText: "{{ Token }}");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# Title", styleSource);

        using var stream = new MemoryStream();
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            () => store.Render(
                new Model
                {
                    Token = new(aroundLandscape)
                },
                stream));

        await Assert.That(exception!.Message).Contains("[SECTION]");
        await Assert.That(exception.PartUri).IsEqualTo("/word/header1.xml");
    }

    // The counterpart of a markdown template with no style source.
    [Test]
    public async Task TemplateWithNoSectionIsRejected()
    {
        using var template = Template(
            "Intro\n\n{{ Body }}\n\nOutro",
            mainPart => mainPart.Document!.Body!.Elements<SectionProperties>().Single().Remove());

        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () =>
            {
                using var doc = await Render(
                    template,
                    new()
                    {
                        Body = aroundLandscape
                    });
            });

        await Assert.That(exception!.Message).Contains("[SECTION]");
        await Assert.That(exception.Message).Contains("w:sectPr");
    }

    static IReadOnlyList<SectionProperties> Sections(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<SectionProperties>()
            .ToList();

    static string DescribeAll(IEnumerable<SectionProperties> sections) =>
        string.Join(" | ", sections.Select(Describe));

    static string Describe(SectionProperties section)
    {
        var size = section.GetFirstChild<PageSize>()!;
        var text = $"{size.Width?.Value}x{size.Height?.Value}";
        if (size.Orient?.Value == PageOrientationValues.Landscape)
        {
            return text + " landscape";
        }

        return text;
    }

    static async Task AssertValid(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static async Task<WordprocessingDocument> Render(Stream template, Model model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    // The builder's one section - a page size, margins and page borders - given a header and a
    // footer as well, since those are what a half-built section loses most visibly.
    static MemoryStream Template(string content, Action<MainDocumentPart>? configure = null, string headerText = "Header")
    {
        var stream = DocxTemplateBuilder.Build(content);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(new Paragraph(new Run(new Text(headerText))));
            var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter");
            footerPart.Footer = new(new Paragraph(new Run(new Text("Footer"))));

            var body = mainPart.Document!.Body!;
            var section = body.Elements<SectionProperties>().Single();
            section.PrependChild(
                new FooterReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdFooter"
                });
            section.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdHeader"
                });

            configure?.Invoke(mainPart);
        }

        stream.Position = 0;
        return stream;
    }

    // {{ Body }} in a first section of the template's own, with no header, and {{ Token }} in the
    // document-final one.
    static MemoryStream TwoSectionTemplate() =>
        Template(
            "First\n\n{{ Body }}\n\nSecond\n\n{{ Token }}\n\nEnd",
            mainPart => mainPart.Document!.Body!.Elements<Paragraph>().ElementAt(1).InsertAfterSelf(
                new Paragraph(
                    new ParagraphProperties(
                        new SectionProperties(
                            new PageSize
                            {
                                Width = 5000,
                                Height = 7000
                            })))));
}
