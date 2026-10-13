

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// An anchor is Parchment's own bookmark. A document that arrives already holding one - a
/// registered or half-processed file supplied again as a template or a style source - must not
/// pass it on to the output, whichever part it sits in and whether or not that part has a token.
/// </summary>
public partial class ArrivedAnchorTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public string Name { get; init; } = "Globex";
    }

    const string arrived = Anchors.Prefix + "arrived";

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task TokenFreeStyleSource(bool dotx)
    {
        using var source = Build("Plain body", "Plain header", "Plain footer", "Plain footnote", "Plain endnote", dotx);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# {{ Name }}", source);

        var output = await Render(store);

        await AssertNoAnchors(output);
        await AssertOwnBookmarksKept(output);
        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var mainPart = doc.MainDocumentPart!;
        await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo("Globex");
        await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo("Plain header");
        await Assert.That(mainPart.FooterParts.Single().Footer!.InnerText).IsEqualTo("Plain footer");
    }

    [Test]
    [Arguments("Header {{ Name }}", "Plain footer", "Plain footnote")]
    [Arguments("Plain header", "Footer {{ Name }}", "Plain footnote")]
    [Arguments("Plain header", "Plain footer", "Footnote {{ Name }}")]
    [Arguments("Header {{ Name }}", "Footer {{ Name }}", "Footnote {{ Name }}")]
    public async Task TokenBearingStyleSource(string header, string footer, string footnote)
    {
        using var source = Build("Plain body", header, footer, footnote, "Plain endnote");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# {{ Name }}", source);

        var output = await Render(store);

        await AssertNoAnchors(output);
        await AssertOwnBookmarksKept(output);
        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var mainPart = doc.MainDocumentPart!;
        await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo(header.Replace("{{ Name }}", "Globex"));
        await Assert.That(mainPart.FooterParts.Single().Footer!.InnerText).IsEqualTo(footer.Replace("{{ Name }}", "Globex"));
        await Assert.That(mainPart.FootnotesPart!.Footnotes!.InnerText).IsEqualTo(footnote.Replace("{{ Name }}", "Globex"));
    }

    [Test]
    [Arguments("Plain body", "Plain header", "Plain footer")]
    [Arguments("Body {{ Name }}", "Plain header", "Plain footer")]
    [Arguments("Plain body", "Header {{ Name }}", "Plain footer")]
    [Arguments("Body {{ Name }}", "Header {{ Name }}", "Footer {{ Name }}")]
    public async Task DocxTemplate(string body, string header, string footer)
    {
        using var template = Build(body, header, footer, "Plain footnote", "Plain endnote");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        // Parts without a token are never read at render, so the snapshot already has to be clean
        // of everything but the anchors this registration wrote.
        var snapshot = ((RegisteredDocxTemplate) store.Resolve(typeof(Model))).Snapshot;
        await Assert.That(AllXml(snapshot)).DoesNotContain(arrived);

        var output = await Render(store);

        await AssertNoAnchors(output);
        await AssertOwnBookmarksKept(output);
        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var mainPart = doc.MainDocumentPart!;
        await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo(body.Replace("{{ Name }}", "Globex"));
        await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo(header.Replace("{{ Name }}", "Globex"));
        await Assert.That(mainPart.FooterParts.Single().Footer!.InnerText).IsEqualTo(footer.Replace("{{ Name }}", "Globex"));
    }

    // The realistic way an anchor arrives: the bytes of one registration reused for another.
    [Test]
    public async Task RegisteredSnapshotReusedAsStyleSourceAndTemplate()
    {
        using var template = Build("Body {{ Name }}", "Header {{ Name }}", "Plain footer", "Footnote {{ Name }}", "Plain endnote", arrivedAnchors: false);
        var first = new TemplateStore();
        first.RegisterDocxTemplate<Model>(template);
        var snapshot = ((RegisteredDocxTemplate) first.Resolve(typeof(Model))).Snapshot;
        await Assert.That(AllXml(snapshot)).Contains(Anchors.Prefix);

        var markdown = new TemplateStore();
        using (var stream = new MemoryStream(snapshot))
        {
            markdown.RegisterMarkdownTemplate<Model>("# {{ Name }}", stream);
        }

        var docx = new TemplateStore();
        using (var stream = new MemoryStream(snapshot))
        {
            docx.RegisterDocxTemplate<Model>(stream);
        }

        foreach (var store in new[] { markdown, docx })
        {
            var output = await Render(store);
            await AssertNoAnchors(output);
            using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
            await Assert.That(doc.MainDocumentPart!.HeaderParts.Single().Header!.InnerText).IsEqualTo("Header Globex");
            await Assert.That(doc.MainDocumentPart.FootnotesPart!.Footnotes!.InnerText).IsEqualTo("Footnote Globex");
        }
    }

    // A style source that needs nothing done is still handed on without being rewritten.
    [Test]
    public async Task StyleSourceWithNothingToStripRendersTheSameAsBefore()
    {
        using var clean = Build("Plain body", "Plain header", "Plain footer", "Plain footnote", "Plain endnote", arrivedAnchors: false);
        var cleanBytes = clean.ToArray();
        var cleanStore = new TemplateStore();
        cleanStore.RegisterMarkdownTemplate<Model>("# {{ Name }}", clean);
        var output = await Render(cleanStore);

        // Never opened for writing, so the footer is the entry as supplied.
        await Assert.That(Entry(output, "word/footer1.xml")).IsEquivalentTo(Entry(cleanBytes, "word/footer1.xml"));
        await AssertNoAnchors(output);
    }

    [Test]
    public async Task StripArrivedPairsEachEndWithItsStart()
    {
        var body = new Body(
            new BookmarkStart { Id = "3", Name = arrived },
            new Paragraph(
                new BookmarkEnd { Id = "3" },
                new BookmarkStart { Id = "3", Name = "kept" },
                new Run(new Text("text")),
                new BookmarkEnd { Id = "3" }),
            new Paragraph(new Run(new Text("no end")), new BookmarkStart { Id = "9", Name = arrived + "2" }));

        await Assert.That(Anchors.StripArrived(body)).IsTrue();

        await Assert.That(body.Descendants<BookmarkStart>().Select(_ => _.Name!.Value!)).IsEquivalentTo(["kept"]);
        await Assert.That(body.Descendants<BookmarkEnd>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<BookmarkEnd>().Single().PreviousSibling()).IsTypeOf<Run>();
        await Assert.That(Anchors.StripArrived(body)).IsFalse();
    }

    static async Task<byte[]> Render(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(new Model(), stream);
        return stream.ToArray();
    }

    static async Task AssertNoAnchors(byte[] package) =>
        await Assert.That(AllXml(package)).DoesNotContain(Anchors.Prefix);

    // The template's own bookmark shares an id with nothing and is not Parchment's, so it stays,
    // start and end, in every part that reaches the output.
    static async Task AssertOwnBookmarksKept(byte[] package)
    {
        using var doc = WordprocessingDocument.Open(new MemoryStream(package), false);
        var mainPart = doc.MainDocumentPart!;
        OpenXmlCompositeElement[] roots =
        [
            mainPart.HeaderParts.Single().Header!,
            mainPart.FooterParts.Single().Footer!,
            mainPart.FootnotesPart!.Footnotes!,
            mainPart.EndnotesPart!.Endnotes!
        ];
        foreach (var root in roots)
        {
            await Assert.That(root.Descendants<BookmarkStart>().Select(_ => _.Name!.Value!)).IsEquivalentTo(["own"]);
            await Assert.That(root.Descendants<BookmarkEnd>().Select(_ => _.Id!.Value!)).IsEquivalentTo(["41"]);
        }
    }

    static string AllXml(byte[] package)
    {
        var builder = new StringBuilder();
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            using var reader = new StreamReader(entry.Open());
            builder.AppendLine(reader.ReadToEnd());
        }

        return builder.ToString();
    }

    static byte[] Entry(byte[] package, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = archive.GetEntry(name)!.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    // Every part holds two paragraphs: the first carries the text (and so any token) with an
    // arrived anchor on it, the second is plain and carries another, beside a bookmark of the
    // template's own. The header also holds one outside any paragraph, where Parchment never puts
    // an anchor but a document that has been through other hands can.
    static MemoryStream Build(string body, string header, string footer, string footnote, string endnote, bool dotx = false, bool arrivedAnchors = true)
    {
        var type = WordprocessingDocumentType.Document;
        if (dotx)
        {
            type = WordprocessingDocumentType.Template;
        }

        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, type))
        {
            var mainPart = doc.AddMainDocumentPart();

            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(Content(header, "h", arrivedAnchors));
            if (arrivedAnchors)
            {
                headerPart.Header.InsertAt(new BookmarkStart { Id = "30", Name = arrived + "-loose" }, 0);
                headerPart.Header.Append(new BookmarkEnd { Id = "30" });
            }

            var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter");
            footerPart.Footer = new(Content(footer, "f", arrivedAnchors));

            var footnotesPart = mainPart.AddNewPart<FootnotesPart>("rIdFootnotes");
            footnotesPart.Footnotes = new(new Footnote(Content(footnote, "fn", arrivedAnchors)) { Id = 1 });

            var endnotesPart = mainPart.AddNewPart<EndnotesPart>("rIdEndnotes");
            endnotesPart.Endnotes = new(new Endnote(Content(endnote, "en", arrivedAnchors)) { Id = 1 });

            var documentBody = new Body(Content(body, "b", arrivedAnchors));
            documentBody.Append(
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
                    }));
            mainPart.Document = new(documentBody);
        }

        stream.Position = 0;
        return stream;
    }

    static IEnumerable<OpenXmlElement> Content(string text, string part, bool arrivedAnchors)
    {
        var first = new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        var second = new Paragraph(
            new BookmarkStart { Id = "41", Name = "own" },
            new BookmarkEnd { Id = "41" });
        if (arrivedAnchors)
        {
            first.InsertAt(new BookmarkEnd { Id = "20" }, 0);
            first.InsertAt(new BookmarkStart { Id = "20", Name = $"{arrived}-{part}-1" }, 0);
            second.InsertAt(new BookmarkEnd { Id = "21" }, 0);
            second.InsertAt(new BookmarkStart { Id = "21", Name = $"{arrived}-{part}-2" }, 0);
        }

        return [first, second];
    }
}
