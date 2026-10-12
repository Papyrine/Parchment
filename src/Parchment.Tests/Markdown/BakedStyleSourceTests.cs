// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A markdown template's registration snapshot is the style source with the body already emptied
/// and the compatibility setting already stamped, so a render no longer parses a body it is about
/// to discard. These pin that the output is what it was when every render did that work itself.
/// </summary>
public partial class BakedStyleSourceTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required string Title;
        public required IReadOnlyList<string> Lines;
    }

    static Doc Model() =>
        new()
        {
            Title = "Annual Report",
            Lines = ["one", "two", "three"]
        };

    const string markdown =
        """
        # {{ Title }}

        [TOC]

        - first
        - second

        [SECTION]{orientation=landscape}

        ## Lines {#lines}

        | Line |
        | --- |
        {% for line in Lines -%}
        | {{ line }} |
        {% endfor %}

        [SECTION]{orientation=portrait}

        See [the lines](#lines).
        """;

    [Test]
    public async Task SnapshotBodyHoldsOnlyTheSection()
    {
        var (bytes, _) = TemplateStore.BakeStyleSource(StyleSource(bodyParagraphs: 200));

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.ChildElements.Count).IsEqualTo(1);
        var section = (SectionProperties)body.ChildElements[0];
        await Assert.That(section.GetFirstChild<HeaderReference>()!.Id!.Value).IsEqualTo("rIdHeader");
        await Assert.That(section.GetFirstChild<PageSize>()!.Width!.Value).IsEqualTo(11906u);
    }

    // A section carried by the last paragraph rather than by the body is still the one preserved,
    // and it is lifted to where a render looks for it.
    [Test]
    public async Task SnapshotLiftsASectionHeldByAParagraph()
    {
        var (bytes, _) = TemplateStore.BakeStyleSource(StyleSource(bodyParagraphs: 5, nestedSection: true));

        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.ChildElements.Count).IsEqualTo(1);
        await Assert.That(body.ChildElements[0] is SectionProperties).IsTrue();
    }

    [Test]
    public async Task SnapshotStampsAnExistingSettingsPartOnce()
    {
        var (bytes, baked) = TemplateStore.BakeStyleSource(StyleSource(bodyParagraphs: 5, settings: true));
        var (again, _) = TemplateStore.BakeStyleSource(bytes);

        await Assert.That(baked).IsTrue();
        await Assert.That(CompatibilityModes(bytes)).IsEqualTo(1);
        await Assert.That(CompatibilityModes(again)).IsEqualTo(1);
    }

    // A settings part created at registration would sit ahead of the numbering part a render adds,
    // where it has always followed it - the same content in a different package order. So it is
    // left for the render to create.
    [Test]
    public async Task SnapshotLeavesAMissingSettingsPartToTheRender()
    {
        var (bytes, baked) = TemplateStore.BakeStyleSource(StyleSource(bodyParagraphs: 5));

        await Assert.That(baked).IsFalse();
        using (var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false))
        {
            await Assert.That(doc.MainDocumentPart!.DocumentSettingsPart).IsNull();
        }

        var rendered = await Render(StyleSource(bodyParagraphs: 5));
        await Assert.That(CompatibilityModes(rendered)).IsEqualTo(1);
        using var output = WordprocessingDocument.Open(new MemoryStream(rendered), false);
        var main = output.MainDocumentPart!;
        await Assert.That(main.GetIdOfPart(main.DocumentSettingsPart!)).IsEqualTo("rIdSettings");
        var entries = EntryNames(rendered);
        await Assert.That(entries.IndexOf("word/numbering.xml")).IsGreaterThanOrEqualTo(0);
        await Assert.That(entries.IndexOf("word/numbering.xml")).IsLessThan(entries.IndexOf("word/settings.xml"));
    }

    // The body is discarded, so how much of it there was cannot reach the output: a style source
    // carrying a sample document renders to the same bytes as one whose body is only the section.
    // That covers the bookmark seed too - the body here holds ids far above anything else, and a
    // seed that counted them would number the produced bookmarks differently.
    [Test]
    [Arguments(false, false)]
    [Arguments(true, false)]
    [Arguments(false, true)]
    [Arguments(true, true)]
    public async Task BodyContentDoesNotReachTheOutput(bool settings, bool nestedSection)
    {
        var full = await Render(StyleSource(bodyParagraphs: 200, settings: settings, nestedSection: nestedSection));
        var empty = await Render(StyleSource(bodyParagraphs: 0, settings: settings));

        await Assert.That(full.AsSpan().SequenceEqual(empty)).IsTrue();
    }

    [Test]
    public async Task ProducedBookmarksAreNumberedAboveTheHeaderOnly()
    {
        var rendered = await Render(StyleSource(bodyParagraphs: 200));

        using var doc = WordprocessingDocument.Open(new MemoryStream(rendered), false);
        var ids = doc.MainDocumentPart!.Document!.Body!
            .Descendants<BookmarkStart>()
            .Select(_ => int.Parse(_.Id!.Value!))
            .ToList();

        // The header's bookmark is 7 and the anchor on its token takes 8; the discarded body's ran
        // from 500.
        await Assert.That(ids.Count).IsGreaterThan(0);
        await Assert.That(ids.Min()).IsEqualTo(9);
        await Assert.That(ids.Max()).IsLessThan(500);
    }

    [Test]
    public async Task HeaderTokensBindAndTheSectionSurvives()
    {
        var rendered = await Render(StyleSource(bodyParagraphs: 200));

        using var doc = WordprocessingDocument.Open(new MemoryStream(rendered), false);
        var main = doc.MainDocumentPart!;
        var body = main.Document!.Body!;
        var header = string.Concat(main.HeaderParts.Single().Header!.Descendants<Text>().Select(_ => _.Text));
        var text = string.Concat(body.Descendants<Text>().Select(_ => _.Text));
        var sections = body.Descendants<SectionProperties>().ToList();

        await Assert.That(header).IsEqualTo("Header: Annual Report");
        await Assert.That(text).DoesNotContain("Sample body paragraph");
        await Assert.That(body.LastChild is SectionProperties).IsTrue();
        await Assert.That(sections.Count).IsEqualTo(3);
        foreach (var section in sections)
        {
            await Assert.That(section.GetFirstChild<HeaderReference>()!.Id!.Value).IsEqualTo("rIdHeader");
        }
    }

    // [SECTION] rewrites the preserved section in place. Every render has to do that to its own
    // copy: one that reached the snapshot would hand the next render a section already rewritten.
    [Test]
    public async Task RendersShareNothing()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, new MemoryStream(StyleSource(bodyParagraphs: 50)));

        var first = await Render(store);
        var second = await Render(store);
        var concurrent = await Task.WhenAll(Enumerable.Range(0, 16).Select(_ => Task.Run(() => Render(store))));

        await Assert.That(first.AsSpan().SequenceEqual(second)).IsTrue();
        foreach (var bytes in concurrent)
        {
            await Assert.That(first.AsSpan().SequenceEqual(bytes)).IsTrue();
        }
    }

    static async Task<byte[]> Render(byte[] styleSource)
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, new MemoryStream(styleSource));
        return await Render(store);
    }

    static async Task<byte[]> Render(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(Model(), stream);
        return stream.ToArray();
    }

    static List<string> EntryNames(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        return archive.Entries.Select(_ => _.FullName).ToList();
    }

    static int CompatibilityModes(byte[] bytes)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return doc.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .Descendants<CompatibilitySetting>()
            .Count(_ => _.Name?.InnerText == "compatibilityMode");
    }

    static byte[] StyleSource(int bodyParagraphs, bool settings = false, bool nestedSection = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            doc.ChangeIdOfPart(mainPart, "rId1");
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            var headerParagraph = BuildParagraph("Header: {{ Title }}");
            headerParagraph.Append(
                new BookmarkStart
                {
                    Id = "7",
                    Name = "inHeader"
                },
                new BookmarkEnd
                {
                    Id = "7"
                });
            headerPart.Header = new(headerParagraph);

            if (settings)
            {
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>("rIdAuthoredSettings");
                settingsPart.Settings = new(
                    new Zoom
                    {
                        Percent = "120"
                    });
            }

            for (var i = 0; i < bodyParagraphs; i++)
            {
                var id = (500 + i).ToString();
                var paragraph = BuildParagraph($"Sample body paragraph {i} with {{{{ Unbound }}}}");
                paragraph.Append(
                    new BookmarkStart
                    {
                        Id = id,
                        Name = $"body{i}"
                    },
                    new BookmarkEnd
                    {
                        Id = id
                    });
                body.Append(paragraph);
            }

            var section = new SectionProperties(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdHeader"
                },
                new PageSize
                {
                    Width = 11906,
                    Height = 16838
                },
                new PageMargin
                {
                    Top = 1440,
                    Right = 1440,
                    Bottom = 1440,
                    Left = 1440,
                    Header = 720,
                    Footer = 720,
                    Gutter = 0
                });
            if (nestedSection)
            {
                body.Append(new Paragraph(new ParagraphProperties(section)));
            }
            else
            {
                body.Append(section);
            }
        }

        return stream.ToArray();
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
