using System.IO.Compression;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A render reads only the parts that carry tokens. Every other header, footer and notes part has
/// to come out as the bytes the registration snapshot holds, and a part that does carry tokens
/// has to bind exactly as before, whichever part it is.
/// </summary>
public partial class TokenFreePartTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public string Name { get; init; } = "Globex";
        public bool Flag { get; init; } = true;
        public IReadOnlyList<Line> Lines { get; init; } = [];
    }

    public class Line
    {
        public required string Name { get; init; }
    }

    static readonly string[] nonBodyEntries =
    [
        "word/header1.xml",
        "word/header2.xml",
        "word/header3.xml",
        "word/footer1.xml",
        "word/footer2.xml",
        "word/footer3.xml",
        "word/footnotes.xml",
        "word/endnotes.xml"
    ];

    // Authored in Word: three headers, three footers, footnotes and endnotes, tokens in the body
    // only. Word writes namespace declarations and attribute order the SDK does not, so this is the
    // template where an untouched part and a parsed-and-rewritten one could differ.
    static MemoryStream WordAuthored() =>
        new(File.ReadAllBytes(Path.Combine(ProjectFiles.ProjectDirectory, "Docx", "WordAuthored.docx")));

    [Test]
    public async Task TokenFreePartsAreTheSnapshotsBytes()
    {
        using var template = WordAuthored();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        var snapshot = ((RegisteredDocxTemplate) store.Resolve(typeof(Model))).Snapshot;

        var output = await Render(
            store,
            new()
            {
                Lines = [new() { Name = "first" }, new() { Name = "second" }]
            });

        foreach (var entry in nonBodyEntries)
        {
            await Assert.That(Entry(output, entry)).IsEquivalentTo(Entry(snapshot, entry));
            // What "no token" rests on: registration anchors token paragraphs only.
            await Assert.That(Text(snapshot, entry)).DoesNotContain(Anchors.Prefix);
        }

        // The body did bind, and its anchors went with it.
        var body = Text(output, "word/document.xml");
        await Assert.That(body).Contains("Globex");
        await Assert.That(body).Contains("second");
        await Assert.That(body).DoesNotContain("{{");
        await Assert.That(body).DoesNotContain(Anchors.Prefix);
        await Assert.That(Text(snapshot, "word/document.xml")).Contains(Anchors.Prefix);
    }

    [Test]
    public async Task RenderIsByteIdenticalAcrossRendersAndRegistrations()
    {
        var first = new TemplateStore();
        using (var template = WordAuthored())
        {
            first.RegisterDocxTemplate<Model>(template);
        }

        var second = new TemplateStore();
        using (var template = WordAuthored())
        {
            second.RegisterDocxTemplate<Model>(template);
        }

        var output = await Render(first, new());

        await Assert.That(await Render(first, new())).IsEquivalentTo(output);
        await Assert.That(await Render(second, new())).IsEquivalentTo(output);
    }

    [Test]
    public async Task TokenOnlyInFooterIsSubstituted()
    {
        using var template = Build(
            body: [Paragraph("Plain body")],
            header: [Paragraph("Plain header")],
            footer: [Paragraph("Prepared for {{ Name }}")],
            footnote: [Paragraph("Plain footnote")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        var snapshot = ((RegisteredDocxTemplate) store.Resolve(typeof(Model))).Snapshot;

        var output = await Render(store, new());

        using (var doc = WordprocessingDocument.Open(new MemoryStream(output), false))
        {
            var mainPart = doc.MainDocumentPart!;
            await Assert.That(mainPart.FooterParts.Single().Footer!.InnerText).IsEqualTo("Prepared for Globex");
            await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo("Plain header");
            await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo("Plain body");
        }

        await Assert.That(Text(output, "word/footer1.xml")).DoesNotContain(Anchors.Prefix);
        foreach (var entry in new[] { "word/document.xml", "word/header1.xml", "word/footnotes.xml" })
        {
            await Assert.That(Entry(output, entry)).IsEquivalentTo(Entry(snapshot, entry));
        }
    }

    [Test]
    public async Task LoopAndConditionalInHeaderRunBesideATokenFreeBody()
    {
        using var template = Build(
            body: [Paragraph("Plain body")],
            header:
            [
                Paragraph("{% for line in Lines %}"),
                Paragraph("{{ line.Name }}"),
                Paragraph("{% endfor %}"),
                Paragraph("{% if Flag %}"),
                Paragraph("shown"),
                Paragraph("{% else %}"),
                Paragraph("hidden"),
                Paragraph("{% endif %}")
            ],
            footer: [Paragraph("Plain footer")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = await Render(
            store,
            new()
            {
                Lines = [new() { Name = "a" }, new() { Name = "b" }]
            });

        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var texts = doc.MainDocumentPart!.HeaderParts.Single().Header!
            .Elements<Paragraph>()
            .Select(_ => _.InnerText);
        await Assert.That(texts).IsEquivalentTo(["a", "b", "shown"]);
    }

    // A value with a newline is the one thing the per-part line break pass exists for, and it only
    // arrives through a substitution — so it can only land in a part that is read.
    [Test]
    public async Task LineBreakInFooterValueIsMaterialized()
    {
        using var template = Build(
            body: [Paragraph("Plain body")],
            header: [Paragraph("Plain header")],
            footer: [Paragraph("{{ Name }}")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = await Render(store, new() { Name = "one\ntwo" });

        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var footer = doc.MainDocumentPart!.FooterParts.Single().Footer!;
        await Assert.That(footer.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(footer.InnerText).IsEqualTo("onetwo");
    }

    // The render's cell repair covers what the render emptied. A cell the template itself left
    // without a paragraph is closed at registration, so it is closed in a part no render reads.
    [Test]
    public async Task OpenCellInTokenFreeHeaderIsClosed()
    {
        using var template = Build(
            body: [Paragraph("{{ Name }}")],
            header: [OpenCell(), new Paragraph()]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = await Render(store, new());

        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var cell = doc.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<TableCell>().Single();
        await Assert.That(cell.LastChild).IsTypeOf<Paragraph>();
        await Assert.That(cell.Elements<Paragraph>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task OpenCellInTokenFreeBodyIsClosed()
    {
        using var template = Build(
            body: [OpenCell(), Paragraph("Plain body")],
            header: [Paragraph("{{ Name }}")]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = await Render(store, new());

        using var doc = WordprocessingDocument.Open(new MemoryStream(output), false);
        var cell = doc.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().Single();
        await Assert.That(cell.LastChild).IsTypeOf<Paragraph>();
        await Assert.That(cell.Elements<Paragraph>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task MarkdownFlowLeavesTokenFreePartsAsTheSnapshotsBytes()
    {
        using var styleSource = Build(
            body: [Paragraph("replaced by the markdown")],
            header: [Paragraph("Plain header")],
            footer: [Paragraph("Prepared for {{ Name }}")],
            footnote: [Paragraph("Plain footnote")]);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# {{ Name }}", styleSource);
        var snapshot = ((RegisteredMarkdownTemplate) store.Resolve(typeof(Model))).Snapshot;

        var output = await Render(store, new());

        using (var doc = WordprocessingDocument.Open(new MemoryStream(output), false))
        {
            var mainPart = doc.MainDocumentPart!;
            await Assert.That(mainPart.FooterParts.Single().Footer!.InnerText).IsEqualTo("Prepared for Globex");
            await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo("Globex");
        }

        await Assert.That(Text(output, "word/footer1.xml")).DoesNotContain(Anchors.Prefix);
        foreach (var entry in new[] { "word/header1.xml", "word/footnotes.xml" })
        {
            await Assert.That(Entry(output, entry)).IsEquivalentTo(Entry(snapshot, entry));
        }
    }

    static async Task<byte[]> Render(TemplateStore store, Model model)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        return stream.ToArray();
    }

    static byte[] Entry(byte[] package, string name)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var stream = archive.GetEntry(name)!.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    static string Text(byte[] package, string name) =>
        Encoding.UTF8.GetString(Entry(package, name));

    static MemoryStream Build(
        OpenXmlElement[] body,
        OpenXmlElement[]? header = null,
        OpenXmlElement[]? footer = null,
        OpenXmlElement[]? footnote = null)
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

            if (footer != null)
            {
                var footerPart = mainPart.AddNewPart<FooterPart>("rIdFooter");
                footerPart.Footer = new(footer);
                section.Append(
                    new FooterReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdFooter"
                    });
            }

            if (footnote != null)
            {
                var footnotesPart = mainPart.AddNewPart<FootnotesPart>("rIdFootnotes");
                footnotesPart.Footnotes = new(
                    new Footnote(footnote)
                    {
                        Id = 1
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

    // A cell with no block content at all, which the schema allows and Word refuses to open.
    static Table OpenCell() =>
        new(
            new TableProperties(
                new TableWidth
                {
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(new GridColumn()),
            new TableRow(new TableCell()));
}
