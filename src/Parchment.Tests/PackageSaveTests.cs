// ReSharper disable PartialTypeWithSinglePart

// A package opened for writing is serialised once, by its Dispose - see "A package is saved by its
// Dispose" in contributing.md. Each test here reads back something only that save could have
// written, at one of the sites that used to save explicitly as well.
public partial class PackageSaveTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required string Title { get; init; }
    }

    static readonly Model model = new()
    {
        Title = "Annual Report"
    };

    const string markdown =
        """
        # {{ Title }}

        Body text.

        # Risks {#risks}

        See the summary on page [](#risks).
        """;

    // Reports one page for every bookmark, and keeps what the document held when it was asked.
    class Resolver : IPageNumberResolver
    {
        public string Seen { get; private set; } = "";

        public Task<IReadOnlyDictionary<string, int>> Resolve(Stream docx, Cancel cancel)
        {
            using var doc = WordprocessingDocument.Open(docx, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            Seen = body.InnerText;
            IReadOnlyDictionary<string, int> pages = body
                .Descendants<BookmarkStart>()
                .Select(_ => _.Name!.Value!)
                .Distinct()
                .ToDictionary(_ => _, _ => 73, StringComparer.Ordinal);
            return Task.FromResult(pages);
        }
    }

    [Test]
    public async Task DocxRenderWritesEveryPartItTouched()
    {
        using var template = BuildWithHeader("Body: {{ Title }}", "Header: {{ Title }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var stream = new MemoryStream();
        await store.Render(
            model,
            stream,
            new WordDocumentProperties
            {
                Title = "From properties"
            });
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart!;
        await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo("Body: Annual Report");
        await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo("Header: Annual Report");
        await Assert.That(doc.PackageProperties.Title).IsEqualTo("From properties");
        await Assert.That(mainPart.Document.Descendants<BookmarkStart>().Any()).IsFalse();
        await Assert.That(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc)).IsEmpty();
    }

    [Test]
    public async Task DocxRegistrationSnapshotHoldsWhatRegistrationBaked()
    {
        using var template = BuildWithHeader("Body: {{ Title }}", "Header: {{ Title }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var snapshot = ((RegisteredDocxTemplate) store.Resolve(typeof(Model))).Snapshot;
        using var doc = WordprocessingDocument.Open(new MemoryStream(snapshot), false);
        var mainPart = doc.MainDocumentPart!;

        // The anchors the scan wrote, in both parts, and the compatibility setting.
        await Assert.That(AnchorCount(mainPart.Document!)).IsEqualTo(1);
        await Assert.That(AnchorCount(mainPart.HeaderParts.Single().Header!)).IsEqualTo(1);
        var compatibility = mainPart.DocumentSettingsPart!.Settings!
            .Descendants<CompatibilitySetting>()
            .Single(_ => _.Name?.Value == CompatSettingNameValues.CompatibilityMode);
        await Assert.That(compatibility.Val?.Value).IsEqualTo("15");
    }

    [Test]
    public async Task MarkdownRenderWritesTheBodyAndTheStyleSourceHeader()
    {
        // The header token is bound by PrepareStyleSource and the sample body is emptied by
        // BakeStyleSource, each into a package of its own.
        using var styleSource = BuildWithHeader("Sample body the markdown replaces", "Header: {{ Title }}");
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# {{ Title }}", styleSource);

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart!;
        await Assert.That(mainPart.Document!.Body!.InnerText).IsEqualTo("Annual Report");
        await Assert.That(mainPart.HeaderParts.Single().Header!.InnerText).IsEqualTo("Header: Annual Report");
        await Assert.That(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc)).IsEmpty();
    }

    [Test]
    public async Task PageNumberResolverIsHandedTheFinishedDocumentAndItsNumbersAreWritten()
    {
        var resolver = new Resolver();
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore
        {
            PageNumbers = resolver
        };
        store.RegisterMarkdownTemplate<Model>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        // The resolver reads the stream between the render's package closing and the one that
        // writes the numbers opening, so the first has to have been saved by then.
        await Assert.That(resolver.Seen).Contains("Annual Report");
        await Assert.That(resolver.Seen).Contains("Body text.");
        await Assert.That(resolver.Seen).DoesNotContain("73");

        using var doc = WordprocessingDocument.Open(stream, false);
        var text = doc.MainDocumentPart!.Document!.Body!.InnerText;
        await Assert.That(text).Contains("Annual Report");
        await Assert.That(text).Contains("73");
    }

    [Test]
    public async Task RendersFromManyThreadsAreByteIdentical()
    {
        using var template = BuildWithHeader("Body: {{ Title }}", "Header: {{ Title }}");
        var docx = new TemplateStore();
        docx.RegisterDocxTemplate<Model>(template);

        using var styleSource = BuildWithHeader("Sample body", "Header: {{ Title }}");
        var markdownStore = new TemplateStore
        {
            PageNumbers = new Resolver()
        };
        markdownStore.RegisterMarkdownTemplate<Model>(markdown, styleSource);

        foreach (var store in new[] { docx, markdownStore })
        {
            var expected = await Render(store);
            var renders = await Task.WhenAll(
                Enumerable.Range(0, 32)
                    .Select(_ => Task.Run(() => Render(store))));
            foreach (var render in renders)
            {
                await Assert.That(render.AsSpan().SequenceEqual(expected)).IsTrue();
            }
        }
    }

    static async Task<byte[]> Render(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        return stream.ToArray();
    }

    static int AnchorCount(OpenXmlElement root) =>
        root.Descendants<BookmarkStart>()
            .Count(_ => _.Name?.Value?.StartsWith(Anchors.Prefix, StringComparison.Ordinal) == true);

    static MemoryStream BuildWithHeader(string bodyText, string headerText)
    {
        var stream = DocxTemplateBuilder.Build(bodyText);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(
                new Paragraph(
                    new Run(
                        new Text(headerText)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        })));
            mainPart.Document!.Body!
                .Elements<SectionProperties>()
                .Single()
                .PrependChild(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    });
        }

        stream.Position = 0;
        return stream;
    }
}
