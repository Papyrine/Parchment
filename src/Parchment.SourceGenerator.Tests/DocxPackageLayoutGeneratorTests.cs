using System.IO.Compression;

// The generator finds a template's parts through the package relationships, as the runtime's SDK
// does, so a part's file name and the namespace it is written in do not decide whether it is
// validated.
public class DocxPackageLayoutGeneratorTests
{
    const string model =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Letter
        {
            public string Title { get; set; } = "";
        }
        """;

    const string transitionalMain = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string strictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    const string transitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    [Test]
    public async Task ConventionalPackage_IsValidatedInEveryPart()
    {
        await Assert.That(Run(BuildDocx())).IsEmpty();
        await Assert.That(Run(BuildDocx(body: "{{ Missing }}"))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(BuildDocx(header: "{{ Missing }}"))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(BuildDocx(footer: "{{ Missing }}"))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(BuildDocx(footnote: "{{ Missing }}"))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(BuildDocx(endnote: "{{ Missing }}"))).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task MainPartUnderAnotherName_IsValidated()
    {
        await Assert.That(Run(RenameMain(BuildDocx()))).IsEmpty();
        await Assert.That(Run(RenameMain(BuildDocx(body: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(RenameMain(BuildDocx(header: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(RenameMain(BuildDocx(endnote: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
    }

    // The parts keep their names, and the main part moves to another folder, so every target in
    // its relationships has to climb out of that folder to reach them.
    [Test]
    public async Task MainPartInAnotherFolder_ResolvesRelativeTargets()
    {
        static byte[] Move(byte[] docx) =>
            Repack(
                docx,
                name =>
                {
                    if (name == "word/document.xml")
                    {
                        return "body/main.xml";
                    }

                    if (name == "word/_rels/document.xml.rels")
                    {
                        return "body/_rels/main.xml.rels";
                    }

                    return name;
                },
                (name, text) =>
                {
                    if (name == "word/_rels/document.xml.rels")
                    {
                        return text.Replace("Target=\"/word/", "Target=\"../word/");
                    }

                    return text.Replace("word/document.xml", "body/main.xml");
                });

        await Assert.That(Run(Move(BuildDocx()))).IsEmpty();
        await Assert.That(Run(Move(BuildDocx(body: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(Move(BuildDocx(footer: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(Move(BuildDocx(footnote: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
    }

    // The SDK writes rooted targets ("/word/header1.xml"); Word writes them relative to the part
    // that owns the relationship ("header1.xml").
    [Test]
    public async Task TargetsRelativeToTheSourcePart_AreResolved()
    {
        static byte[] Relative(byte[] docx) =>
            Repack(
                docx,
                _ => _,
                (name, text) =>
                {
                    if (name == "word/_rels/document.xml.rels")
                    {
                        return text.Replace("Target=\"/word/", "Target=\"");
                    }

                    if (name == "_rels/.rels")
                    {
                        return text.Replace("Target=\"/word/", "Target=\"word/");
                    }

                    return text;
                });

        await Assert.That(Run(Relative(BuildDocx()))).IsEmpty();
        await Assert.That(Run(Relative(BuildDocx(body: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(Relative(BuildDocx(header: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
    }

    // The runtime registers and renders a Strict Open XML document, so the generator validates it
    // the same as a transitional one, and finds its settings part.
    [Test]
    public async Task StrictDocument_IsValidatedInEveryPart()
    {
        await Assert.That(Run(ToStrict(BuildDocx()))).IsEmpty();
        await Assert.That(Run(ToStrict(BuildDocx(body: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(ToStrict(BuildDocx(header: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(ToStrict(BuildDocx(footer: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(ToStrict(BuildDocx(footnote: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Run(ToStrict(BuildDocx(endnote: "{{ Missing }}")))).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task StrictDocumentWithoutTheSetting_ReportsParch012()
    {
        var docx = ToStrict(BuildDocx(removePersonalInformation: false));

        await Assert.That(Run(docx)).IsEquivalentTo(["PARCH012"]);
    }

    // The runtime never opens a part the main part has no relationship to, whatever its name.
    [Test]
    public async Task PartNoRelationshipPointsAt_IsNotRead()
    {
        var docx = Repack(
            BuildDocx(header: "{{ Missing }}"),
            _ => _,
            (name, text) =>
            {
                if (name == "word/_rels/document.xml.rels")
                {
                    return Regex.Replace(text, "<Relationship [^>]*header[^>]*/>", "");
                }

                return text;
            });

        await Assert.That(Run(docx)).IsEmpty();
    }

    [Test]
    public async Task ZipWithoutAWordPart_ReportsParch006()
    {
        using var stream = new MemoryStream();
        await using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            await using var writer = new StreamWriter(await archive.CreateEntry("readme.txt").OpenAsync());
            await writer.WriteAsync("not a document");
        }

        await Assert.That(Run(stream.ToArray())).IsEquivalentTo(["PARCH006"]);
    }

    [Test]
    public async Task MainPartMissingFromThePackage_ReportsParch006()
    {
        var docx = Repack(
            BuildDocx(),
            name =>
            {
                if (name == "word/document.xml")
                {
                    return "word/orphan.xml";
                }

                return name;
            },
            (_, text) => text);

        await Assert.That(Run(docx)).IsEquivalentTo(["PARCH006"]);
    }

    // Another kind of Office package: it has a main part, and that part is not a Word document.
    [Test]
    public async Task MainPartThatIsNotAWordDocument_ReportsParch006()
    {
        var docx = Repack(
            BuildDocx(),
            _ => _,
            (name, text) =>
            {
                if (name == "word/document.xml")
                {
                    return """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>""";
                }

                return text;
            });

        await Assert.That(Run(docx)).IsEquivalentTo(["PARCH006"]);
    }

    static List<string> Run(byte[] docx)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            model,
            new TemplateFile("Letter.parchment.docx", docx));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
    }

    static byte[] RenameMain(byte[] docx) =>
        Repack(
            docx,
            name => name
                .Replace("word/document.xml", "word/document2.xml")
                .Replace("word/_rels/document.xml.rels", "word/_rels/document2.xml.rels"),
            (_, text) => text.Replace("word/document.xml", "word/document2.xml"));

    static byte[] ToStrict(byte[] docx) =>
        Repack(
            docx,
            _ => _,
            (_, text) => text
                .Replace(transitionalMain, strictMain)
                .Replace(transitionalRelationships, strictRelationships));

    static byte[] Repack(byte[] docx, Func<string, string> rename, Func<string, string, string> rewrite)
    {
        using var input = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        using var stream = new MemoryStream();
        using (var output = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            foreach (var entry in input.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                var text = rewrite(entry.FullName, reader.ReadToEnd());
                using var writer = new StreamWriter(output.CreateEntry(rename(entry.FullName)).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
        }

        return stream.ToArray();
    }

    static byte[] BuildDocx(
        string body = "{{ Title }}",
        string header = "{{ Title }}",
        string footer = "{{ Title }}",
        string footnote = "{{ Title }}",
        string endnote = "{{ Title }}",
        bool removePersonalInformation = true)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(TextParagraph(body)));

            mainPart.AddNewPart<HeaderPart>().Header = new(TextParagraph(header));
            mainPart.AddNewPart<FooterPart>().Footer = new(TextParagraph(footer));
            mainPart.AddNewPart<FootnotesPart>().Footnotes = new(
                new Footnote(TextParagraph(footnote))
                {
                    Id = 1
                });
            mainPart.AddNewPart<EndnotesPart>().Endnotes = new(
                new Endnote(TextParagraph(endnote))
                {
                    Id = 1
                });

            var settings = new Settings();
            if (removePersonalInformation)
            {
                settings.Append(new RemovePersonalInformation(), new RemoveDateAndTime());
            }

            mainPart.AddNewPart<DocumentSettingsPart>().Settings = settings;
        }

        return stream.ToArray();
    }

    static Paragraph TextParagraph(string text) =>
        new(new Run(new Text(text)));
}
