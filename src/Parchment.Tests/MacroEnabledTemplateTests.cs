// ReSharper disable PartialTypeWithSinglePart
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Validation;

// A macro-enabled template (.docm/.dotm) is retyped to a plain document like any other, so the VBA
// project it carries has to go with the type. Left in, the output is a .docx-typed package holding
// vbaProject.bin: a mislabelled file that Word and security tooling both treat as suspect.
public partial class MacroEnabledTemplateTests
{
    [ParchmentBindable]
    public partial class MacroModel
    {
        public required string Title { get; init; }
    }

    [Test]
    [Arguments(WordprocessingDocumentType.MacroEnabledDocument)]
    [Arguments(WordprocessingDocumentType.MacroEnabledTemplate)]
    // Already mislabelled on the way in: typed as a document, VBA project inside.
    [Arguments(WordprocessingDocumentType.Document)]
    public async Task DocxTemplateRendersWithoutMacros(WordprocessingDocumentType type)
    {
        using var template = BuildMacroEnabled(type, "{{ Title }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<MacroModel>(template);

        using var output = await Render(store);

        await AssertMacroFree(output);
        using var result = WordprocessingDocument.Open(output, false);
        await Assert.That(result.MainDocumentPart!.Document!.Body!.InnerText).IsEqualTo("x");
    }

    [Test]
    [Arguments(WordprocessingDocumentType.MacroEnabledDocument)]
    [Arguments(WordprocessingDocumentType.MacroEnabledTemplate)]
    [Arguments(WordprocessingDocumentType.Document)]
    public async Task MarkdownStyleSourceRendersWithoutMacros(WordprocessingDocumentType type)
    {
        using var styleSource = BuildMacroEnabled(type, "style source body");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<MacroModel>("# {{ Title }}", styleSource);

        using var output = await Render(store);

        await AssertMacroFree(output);
        using var result = WordprocessingDocument.Open(output, false);
        await Assert.That(result.MainDocumentPart!.Document!.Body!.InnerText).IsEqualTo("x");
    }

    // The token in the header makes registration rewrite the style source for a second reason, so
    // this is the path where the macro parts are removed and the package is saved with scope anchors.
    [Test]
    public async Task MarkdownStyleSourceWithHeaderTokenRendersWithoutMacros()
    {
        using var styleSource = BuildMacroEnabled(WordprocessingDocumentType.MacroEnabledTemplate, "body", "{{ Title }}");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<MacroModel>("# {{ Title }}", styleSource);

        using var output = await Render(store);

        await AssertMacroFree(output);
        using var result = WordprocessingDocument.Open(output, false);
        await Assert.That(result.MainDocumentPart!.HeaderParts.Single().Header!.InnerText).IsEqualTo("x");
    }

    // The content types are rewritten outside the packaging API, so this is the path to hold to
    // the same-bytes guarantee.
    [Test]
    public async Task RenderIsByteIdentical()
    {
        // Built once: the fixture's relationship ids are minted at random.
        using var fixture = BuildMacroEnabled(WordprocessingDocumentType.MacroEnabledDocument, "{{ Title }}");
        var bytes = new List<byte[]>();
        for (var i = 0; i < 2; i++)
        {
            using var template = new MemoryStream(fixture.ToArray());
            var store = new TemplateStore();
            store.RegisterDocxTemplate<MacroModel>(template);
            using var output = await Render(store);
            bytes.Add(output.ToArray());
        }

        await Assert.That(bytes[0].SequenceEqual(bytes[1])).IsTrue();
    }

    static async Task<MemoryStream> Render(TemplateStore store)
    {
        var output = new MemoryStream();
        await store.Render(
            new MacroModel
            {
                Title = "x"
            },
            output);
        output.Position = 0;
        return output;
    }

    static async Task AssertMacroFree(MemoryStream output)
    {
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            await Assert.That(doc.DocumentType).IsEqualTo(WordprocessingDocumentType.Document);

            var main = doc.MainDocumentPart!;
            await Assert.That(main.VbaProjectPart).IsNull();
            await Assert.That(main.CustomizationPart).IsNull();

            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
                .Validate(doc)
                .Select(_ => _.Description)
                .ToList();
            await Assert.That(errors).IsEmpty();
        }

        output.Position = 0;

        // The package as a consumer's tooling sees it: no macro entry, no macro content type - not
        // even an unused default - and a content type for every part that is there.
        using var zip = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
        var entries = zip.Entries
            .Select(_ => "/" + _.FullName)
            .ToList();
        var macroEntries = entries
            .Where(_ => _.Contains("vba", StringComparison.OrdinalIgnoreCase) ||
                        _.Contains("customizations", StringComparison.OrdinalIgnoreCase) ||
                        _.Contains("attachedToolbars", StringComparison.OrdinalIgnoreCase))
            .ToList();
        await Assert.That(macroEntries).IsEmpty();

        XDocument types;
        await using (var typesStream = zip.GetEntry("[Content_Types].xml")!.Open())
        {
            types = XDocument.Load(typesStream);
        }

        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
        var defaults = types.Root!
            .Elements(ns + "Default")
            .ToDictionary(_ => _.Attribute("Extension")!.Value, _ => _.Attribute("ContentType")!.Value, StringComparer.OrdinalIgnoreCase);
        var overrides = types.Root
            .Elements(ns + "Override")
            .ToDictionary(_ => _.Attribute("PartName")!.Value, _ => _.Attribute("ContentType")!.Value, StringComparer.OrdinalIgnoreCase);

        var macroTypes = defaults.Values
            .Concat(overrides.Values)
            .Where(_ => _.Contains("vba", StringComparison.OrdinalIgnoreCase) ||
                        _.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase) ||
                        _.Contains("keyMapCustomizations", StringComparison.OrdinalIgnoreCase) ||
                        _.Contains("attachedToolbars", StringComparison.OrdinalIgnoreCase))
            .ToList();
        await Assert.That(macroTypes).IsEmpty();

        var missing = overrides.Keys
            .Where(_ => !entries.Contains(_))
            .ToList();
        await Assert.That(missing).IsEmpty();

        var untyped = entries
            .Where(_ => _ != "/[Content_Types].xml" &&
                        !overrides.ContainsKey(_) &&
                        !defaults.ContainsKey(Path.GetExtension(_).TrimStart('.')))
            .ToList();
        await Assert.That(untyped).IsEmpty();

        output.Position = 0;
    }

    // What Word writes for a macro-enabled file: the VBA project off the main part with its
    // supplementary data beneath it, and the key and toolbar customizations that point into it.
    static MemoryStream BuildMacroEnabled(WordprocessingDocumentType type, string body, string? header = null)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, type))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(new Paragraph(new Run(new Text(body)))));

            if (header != null)
            {
                var headerPart = main.AddNewPart<HeaderPart>();
                headerPart.Header = new(new Paragraph(new Run(new Text(header))));
                main.Document.Body!.AppendChild(
                    new SectionProperties(
                        new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = main.GetIdOfPart(headerPart)
                        }));
            }

            var vba = main.AddNewPart<VbaProjectPart>();
            Feed(vba, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
            Feed(
                vba.AddNewPart<VbaDataPart>(),
                """<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>"""u8.ToArray());

            var customization = main.AddNewPart<CustomizationPart>();
            Feed(
                customization,
                """<wne:tcg xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>"""u8.ToArray());
            Feed(customization.AddNewPart<WordAttachedToolbarsPart>(), [1, 2, 3, 4]);
        }

        stream.Position = 0;
        return stream;
    }

    static void Feed(OpenXmlPart part, byte[] bytes)
    {
        using var data = new MemoryStream(bytes);
        part.FeedData(data);
    }
}
