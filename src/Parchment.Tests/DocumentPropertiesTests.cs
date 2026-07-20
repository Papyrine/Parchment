using System.Xml.Linq;
using CustomProps = DocumentFormat.OpenXml.CustomProperties;

// Excelsior exposes a DocumentProperties of its own and this project imports both namespaces.
using DocumentProperties = Parchment.DocumentProperties;

public class DocumentPropertiesTests
{
    public class Model
    {
        public required string Title { get; init; }
    }

    // A template normally arrives carrying properties of its own. Rewriting docProps/custom.xml
    // wholesale drops them without a word — the trap both hand-rolled stampers this replaces hit.
    static MemoryStream BuildTemplateWithOwnProperties()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(new Paragraph(new Run(new Text("{{ Title }}")))));

            var custom = doc.AddCustomFilePropertiesPart();
            custom.Properties = new(
                new CustomProps.CustomDocumentProperty(
                    new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR("legislation;bills"))
                {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    PropertyId = 2,
                    Name = "ESearchTags"
                });

            var core = doc.AddCoreFilePropertiesPart();
            using var coreStream = core.GetStream(FileMode.Create);
            new XDocument(
                new XElement(
                    XNamespace.Get("http://schemas.openxmlformats.org/package/2006/metadata/core-properties") + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"),
                    new XAttribute(XNamespace.Xmlns + "dc", "http://purl.org/dc/elements/1.1/"),
                    new XElement(XNamespace.Get("http://purl.org/dc/elements/1.1/") + "creator", "Original Author"))).Save(coreStream);
        }

        stream.Position = 0;
        return stream;
    }

    static async Task<WordprocessingDocument> Render(DocumentProperties? properties)
    {
        using var template = BuildTemplateWithOwnProperties();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>("t", template);

        var output = new MemoryStream();
        await store.Render(
            "t",
            new Model
            {
                Title = "x"
            },
            output,
            properties);
        output.Position = 0;
        return WordprocessingDocument.Open(output, false);
    }

    static string? Custom(WordprocessingDocument doc, string name) =>
        doc.CustomFilePropertiesPart?.Properties
            ?.Elements<CustomProps.CustomDocumentProperty>()
            .FirstOrDefault(_ => _.Name?.Value == name)
            ?.InnerText;

    static string? Core(WordprocessingDocument doc, XName name)
    {
        using var stream = doc.CoreFilePropertiesPart!.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream).Root!.Element(name)?.Value;
    }

    static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";
    static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    [Test]
    public async Task CoreValuesAreWritten()
    {
        using var doc = await Render(
            new()
            {
                Title = "Bill 42",
                Subject = "Second reading",
                Status = "Final"
            });

        await Assert.That(Core(doc, dc + "title")).IsEqualTo("Bill 42");
        await Assert.That(Core(doc, dc + "subject")).IsEqualTo("Second reading");
        await Assert.That(Core(doc, cp + "contentStatus")).IsEqualTo("Final");
    }

    // Only the values set are touched, so a core property the template carries survives.
    [Test]
    public async Task UnsetCoreValuesLeaveTheTemplatesOwnAlone()
    {
        using var doc = await Render(
            new()
            {
                Title = "Bill 42"
            });

        await Assert.That(Core(doc, dc + "creator")).IsEqualTo("Original Author");
    }

    [Test]
    public async Task SetCoreValueReplacesTheTemplatesOwn()
    {
        using var doc = await Render(
            new()
            {
                Author = "New Author"
            });

        await Assert.That(Core(doc, dc + "creator")).IsEqualTo("New Author");
    }

    // The sharp edge: a template's own custom properties must survive a stamp.
    [Test]
    public async Task CustomPropertiesAreMergedNotReplaced()
    {
        using var doc = await Render(
            new()
            {
                Custom =
                {
                    ["BillNumber"] = "42"
                }
            });

        await Assert.That(Custom(doc, "ESearchTags")).IsEqualTo("legislation;bills");
        await Assert.That(Custom(doc, "BillNumber")).IsEqualTo("42");
    }

    [Test]
    public async Task CustomPropertyWithSameNameReplacesTheTemplatesOwn()
    {
        using var doc = await Render(
            new()
            {
                Custom =
                {
                    ["ESearchTags"] = "replaced"
                }
            });

        await Assert.That(Custom(doc, "ESearchTags")).IsEqualTo("replaced");

        var count = doc.CustomFilePropertiesPart!.Properties!
            .Elements<CustomProps.CustomDocumentProperty>()
            .Count(_ => _.Name?.Value == "ESearchTags");
        await Assert.That(count).IsEqualTo(1);
    }

    [Test]
    public async Task RemoveCustomDropsTheTemplatesOwn()
    {
        using var doc = await Render(
            new()
            {
                RemoveCustom = { "ESearchTags" }
            });

        await Assert.That(Custom(doc, "ESearchTags")).IsNull();
    }

    // PropertyId is 1-based, 1 is reserved, and the ids must stay contiguous across edits.
    [Test]
    public async Task CustomPropertyIdsStayContiguousFromTwo()
    {
        using var doc = await Render(
            new()
            {
                Custom =
                {
                    ["A"] = "1",
                    ["B"] = "2"
                }
            });

        var ids = doc.CustomFilePropertiesPart!.Properties!
            .Elements<CustomProps.CustomDocumentProperty>()
            .Select(_ => _.PropertyId!.Value)
            .ToList();
        await Assert.That(ids).IsEquivalentTo([2, 3, 4]);
    }

    [Test]
    public async Task ExtendedValuesAreWritten()
    {
        using var doc = await Render(
            new()
            {
                Company = "Papyrine",
                Manager = "Alex"
            });

        await Assert.That(doc.ExtendedFilePropertiesPart!.Properties!.Company!.Text).IsEqualTo("Papyrine");
        await Assert.That(doc.ExtendedFilePropertiesPart.Properties.Manager!.Text).IsEqualTo("Alex");
    }

    // Coercing an unsupported type would write something like "System.Int32[]" into the property
    // and hide the mistake.
    [Test]
    public async Task UnsupportedCustomValueTypeThrows()
    {
        var exception = await Assert.ThrowsAsync<ArgumentException>(
            async () =>
            {
                using var _ = await Render(
                    new()
                    {
                        Custom =
                        {
                            ["Bad"] = new[] { 1, 2 }
                        }
                    });
            });
        await Assert.That(exception!.Message).Contains("Bad");
    }

    [Test]
    public async Task NoPropertiesLeavesTheTemplateUntouched()
    {
        using var doc = await Render(null);

        await Assert.That(Custom(doc, "ESearchTags")).IsEqualTo("legislation;bills");
        await Assert.That(Core(doc, dc + "creator")).IsEqualTo("Original Author");
    }

    // Every other test here registers through RegisterDocxTemplate, so the markdown flow's call to
    // DocumentPropertiesWriter.Apply was wired but unproven. The writer is shared, so this covers
    // the wiring rather than the merge semantics: one value written, one the style source carries
    // left alone.
    [Test]
    public async Task MarkdownFlowWritesAndMergesProperties()
    {
        using var styleSource = BuildTemplateWithOwnProperties();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("md", "# {{ Title }}", styleSource);

        using var output = new MemoryStream();
        await store.Render(
            "md",
            new Model
            {
                Title = "x"
            },
            output,
            new DocumentProperties
            {
                Title = "Bill 42",
                Custom =
                {
                    ["Chamber"] = "Lower"
                }
            });
        output.Position = 0;

        using var doc = WordprocessingDocument.Open(output, false);
        await Assert.That(Core(doc, dc + "title")).IsEqualTo("Bill 42");
        await Assert.That(Custom(doc, "Chamber")).IsEqualTo("Lower");
        await Assert.That(Custom(doc, "ESearchTags")).IsEqualTo("legislation;bills");
    }
}
