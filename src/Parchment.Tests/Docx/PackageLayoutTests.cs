using System.IO.Compression;
using DocumentFormat.OpenXml.Validation;

// Templates the source generator has to agree with the runtime on: a main part that is not named
// word/document.xml, and a document saved as Strict Open XML. Both register and render, so the
// generator validates both (DocxPackageLayoutGeneratorTests).
public partial class PackageLayoutTests
{
    [ParchmentBindable]
    public partial class Letter
    {
        public string Name { get; set; } = "World";
    }

    [Test]
    public async Task MainPartUnderAnotherName_Renders()
    {
        using var template = DocxTemplateBuilder.Build("Hello {{ Name }}");
        using var moved = Repack(
            template,
            name => name
                .Replace("word/document.xml", "word/document2.xml")
                .Replace("word/_rels/document.xml.rels", "word/_rels/document2.xml.rels"),
            text => text.Replace("word/document.xml", "word/document2.xml"));

        await AssertRenders(moved);
    }

    [Test]
    public async Task StrictDocument_Renders()
    {
        using var template = DocxTemplateBuilder.Build("Hello {{ Name }}");
        using var strict = ToStrict(template);

        await AssertRenders(strict);
    }

    [Test]
    public async Task StrictDocument_MissingMemberFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build("Hello {{ Missing }}");
        using var strict = ToStrict(template);
        var store = new TemplateStore();

        await Assert.That(() => store.RegisterDocxTemplate<Letter>(strict))
            .Throws<ParchmentRegistrationException>()
            .WithMessageContaining("Identifier 'Missing'");
    }

    static async Task AssertRenders(Stream template)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Letter>(template);
        using var output = new MemoryStream();
        await store.Render(new Letter(), output);
        output.Position = 0;

        using var document = WordprocessingDocument.Open(output, false);
        await Assert.That(document.MainDocumentPart!.Document!.Body!.InnerText).IsEqualTo("Hello World");
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .Select(_ => _.Description)
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static MemoryStream ToStrict(Stream docx) =>
        Repack(
            docx,
            _ => _,
            text => text
                .Replace("<w:document ", "<w:document w:conformance=\"strict\" ")
                .Replace("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main")
                .Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"));

    static MemoryStream Repack(Stream docx, Func<string, string> rename, Func<string, string> rewrite)
    {
        using var input = new ZipArchive(docx, ZipArchiveMode.Read, true);
        var stream = new MemoryStream();
        using (var output = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            foreach (var entry in input.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                var text = rewrite(reader.ReadToEnd());
                using var writer = new StreamWriter(output.CreateEntry(rename(entry.FullName)).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
        }

        stream.Position = 0;
        return stream;
    }
}
