using System.Xml.Linq;
using CustomProps = DocumentFormat.OpenXml.CustomProperties;
using ExtendedProps = DocumentFormat.OpenXml.ExtendedProperties;

// Excelsior exposes a DocumentProperties of its own and this project imports both namespaces.
// begin-snippet: DocumentPropertiesAlias
using DocumentProperties = Parchment.DocumentProperties;
// end-snippet

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
                    new XElement(XNamespace.Get("http://purl.org/dc/elements/1.1/") + "creator", "Original Author"),
                    // Editing history this type cannot set, and the reason clearing exists at all:
                    // it describes work on the template, not on the generated document.
                    new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/metadata/core-properties") + "revision", "21"),
                    new XElement(XNamespace.Get("http://schemas.openxmlformats.org/package/2006/metadata/core-properties") + "lastPrinted", "2023-03-30T04:27:00Z"))).Save(coreStream);

            var extended = doc.AddExtendedFilePropertiesPart();
            extended.Properties = new(
                new ExtendedProps.Company("Original Company"),
                new ExtendedProps.Manager("Original Manager"));
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

    // The three parts in one pass, as the readme shows it.
    [Test]
    public async Task ValuesLandAcrossAllThreeParts()
    {
        using var template = BuildTemplateWithOwnProperties();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>("bill", template);
        var model = new Model
        {
            Title = "x"
        };
        using var stream = new MemoryStream();

        #region DocumentProperties

        await store.Render(
            "bill",
            model,
            stream,
            new DocumentProperties
            {
                Title = "Bill 42",
                Author = "Drafting Office",
                Status = "Final",
                Company = "Papyrine",
                Custom =
                {
                    ["BillNumber"] = "42",
                    ["Introduced"] = new Date(2026, 3, 1)
                }
            });

        #endregion

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Core(doc, dc + "title")).IsEqualTo("Bill 42");
        await Assert.That(Core(doc, dc + "creator")).IsEqualTo("Drafting Office");
        await Assert.That(Core(doc, cp + "contentStatus")).IsEqualTo("Final");
        await Assert.That(doc.ExtendedFilePropertiesPart!.Properties!.Company!.Text).IsEqualTo("Papyrine");
        await Assert.That(Custom(doc, "BillNumber")).IsEqualTo("42");
    }

    // A template someone has edited names them, and that name should not travel with every document
    // generated from it.
    [Test]
    public async Task ClearBuiltInDropsWhatTheTemplateCarried()
    {
        using var doc = await Render(
            new()
            {
                ClearBuiltIn = true
            });

        await Assert.That(Core(doc, dc + "creator")).IsNull();
    }

    // The part holds more than this type can set, and those are the values describing the
    // template's own history, so clearing has to reach them too.
    [Test]
    public async Task ClearBuiltInDropsValuesThatCannotBeSet()
    {
        using var doc = await Render(
            new()
            {
                ClearBuiltIn = true
            });

        await Assert.That(Core(doc, cp + "revision")).IsNull();
        await Assert.That(Core(doc, cp + "lastPrinted")).IsNull();
    }

    // Clearing happens first, so the values the caller did set still land.
    [Test]
    public async Task ClearBuiltInStillWritesTheValuesSet()
    {
        #region ClearBuiltIn

        var properties = new DocumentProperties
        {
            ClearBuiltIn = true,
            Title = "Bill 42"
        };

        #endregion

        using var doc = await Render(properties);

        await Assert.That(Core(doc, dc + "title")).IsEqualTo("Bill 42");
        await Assert.That(Core(doc, dc + "creator")).IsNull();
    }

    [Test]
    public async Task ClearBuiltInDropsCompanyAndManager()
    {
        using var doc = await Render(
            new()
            {
                ClearBuiltIn = true
            });

        var extended = doc.ExtendedFilePropertiesPart!.Properties!;
        await Assert.That(extended.Company?.Text).IsNull();
        await Assert.That(extended.Manager?.Text).IsNull();
    }

    // User-defined properties are the template's own data rather than metadata about who edited it,
    // so they are left alone — RemoveCustom drops those by name.
    [Test]
    public async Task ClearBuiltInLeavesCustomPropertiesAlone()
    {
        using var doc = await Render(
            new()
            {
                ClearBuiltIn = true
            });

        await Assert.That(Custom(doc, "ESearchTags")).IsEqualTo("legislation;bills");
    }
}
