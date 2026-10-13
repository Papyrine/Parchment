using Parchment.Generated;

// Input that cannot be a template fails at registration, as a ParchmentRegistrationException naming
// the template. It used to surface as a NullReferenceException or a raw framework exception, and an
// unusable markdown style source registered and then failed at the first render.
public partial class UnusableTemplateTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public string Title { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class GeneratedDocxModel
    {
        public string Title { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class GeneratedMarkdownModel
    {
        public string Title { get; set; } = "";
    }

    sealed class WriteOnlyStream : MemoryStream
    {
        public override bool CanRead => false;
    }

    sealed class NonSeekableStream(byte[] bytes) : MemoryStream(bytes)
    {
        public override bool CanSeek => false;
    }

    static byte[] Zip(Action<ZipArchive> fill)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            fill(zip);
        }

        return stream.ToArray();
    }

    static byte[] Word(Action<WordprocessingDocument> fill)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            fill(doc);
        }

        return stream.ToArray();
    }

    static byte[] Valid()
    {
        using var stream = DocxTemplateBuilder.Build("{{ Title }}");
        return stream.ToArray();
    }

    static byte[] MalformedXml()
    {
        using var stream = new MemoryStream();
        stream.Write(Valid());
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true))
        {
            zip.GetEntry("word/document.xml")!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open());
            writer.Write("""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p></w:body>""");
        }

        return stream.ToArray();
    }

    static byte[] Spreadsheet()
    {
        using var stream = new MemoryStream();
        using (var sheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            sheet.AddWorkbookPart().Workbook = new();
        }

        return stream.ToArray();
    }

    static byte[] Presentation()
    {
        using var stream = new MemoryStream();
        using (var deck = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            deck.AddPresentationPart().Presentation = new();
        }

        return stream.ToArray();
    }

    static byte[] CompoundFile()
    {
        var bytes = new byte[4096];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(bytes, 0);
        return bytes;
    }

    // Each unusable input, what its message has to say, and whether a framework exception is behind it.
    static IEnumerable<(string Label, byte[] Bytes, string Expected, bool HasInner)> Cases()
    {
        var valid = Valid();
        yield return ("empty", [], "empty", false);
        yield return ("random bytes", [.. Enumerable.Range(0, 2000).Select(_ => (byte) (_ * 7))], "Open XML package", true);
        yield return ("plain text", "hello, not a docx"u8.ToArray(), "Open XML package", true);
        yield return ("compound file", CompoundFile(), "password-protected", true);
        yield return ("truncated", valid[..(valid.Length / 2)], "Open XML package", true);
        yield return ("zip that is not a package", Zip(_ => _.CreateEntry("hello.txt")), "no main document part", false);
        yield return ("no main part", Word(_ => { }), "no main document part", false);
        yield return ("no root", Word(_ => _.AddMainDocumentPart()), "no document body", false);
        yield return ("no body", Word(_ => _.AddMainDocumentPart().Document = new()), "no document body", false);
        yield return ("malformed xml", MalformedXml(), "malformed", true);
        yield return ("spreadsheet", Spreadsheet(), "spreadsheet", true);
        yield return ("presentation", Presentation(), "presentation", true);
    }

    static async Task AssertRejected(string label, string subject, string expected, bool hasInner, Action register)
    {
        ParchmentRegistrationException exception;
        try
        {
            register();
            throw new($"{label}: registered, but should have been rejected");
        }
        catch (ParchmentRegistrationException caught)
        {
            exception = caught;
        }
        catch (Exception other) when (other is not ParchmentException)
        {
            throw new($"{label}: threw {other.GetType().Name} rather than ParchmentRegistrationException", other);
        }

        await Assert.That(exception.TemplateName).IsEqualTo(nameof(Model));
        await Assert.That(exception.Message).Contains($"The {subject}").Because(label);
        await Assert.That(exception.Message).Contains(expected).Because(label);
        await Assert.That(exception.InnerException != null).IsEqualTo(hasInner).Because(label);
    }

    [Test]
    public async Task DocxTemplate()
    {
        foreach (var (label, bytes, expected, hasInner) in Cases())
        {
            var subject = "template";
            if (hasInner && expected is "malformed" or "spreadsheet" or "presentation")
            {
                subject = "main document part of the template";
            }

            await AssertRejected(
                label,
                subject,
                expected,
                hasInner,
                () =>
                {
                    using var stream = new MemoryStream(bytes);
                    new TemplateStore().RegisterDocxTemplate<Model>(stream);
                });
        }
    }

    [Test]
    public async Task MarkdownStyleSource()
    {
        foreach (var (label, bytes, expected, hasInner) in Cases())
        {
            var subject = "style source";
            if (hasInner && expected is "malformed" or "spreadsheet" or "presentation")
            {
                subject = "main document part of the style source";
            }

            await AssertRejected(
                label,
                subject,
                expected,
                hasInner,
                () =>
                {
                    using var stream = new MemoryStream(bytes);
                    new TemplateStore().RegisterMarkdownTemplate<Model>("# {{ Title }}", stream);
                });
        }
    }

    [Test]
    public async Task EmptyFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"{nameof(UnusableTemplateTests)}.{nameof(EmptyFile)}.docx");
        await File.WriteAllBytesAsync(path, []);
        try
        {
            await AssertRejected("empty file", "template", "empty", false, () => new TemplateStore().RegisterDocxTemplate<Model>(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task UnreadableStreams()
    {
        var disposed = new MemoryStream(Valid());
        disposed.Dispose();

        await AssertRejected("disposed template", "template", "cannot be read", false, () => new TemplateStore().RegisterDocxTemplate<Model>(disposed));
        await AssertRejected("disposed style source", "style source", "cannot be read", false, () => new TemplateStore().RegisterMarkdownTemplate<Model>("# {{ Title }}", disposed));

        using var writeOnly = new WriteOnlyStream();
        await AssertRejected("write-only template", "template", "cannot be read", false, () => new TemplateStore().RegisterDocxTemplate<Model>(writeOnly));
        await AssertRejected("write-only style source", "style source", "cannot be read", false, () => new TemplateStore().RegisterMarkdownTemplate<Model>("# {{ Title }}", writeOnly));
    }

    // A null argument is a mistake in the calling code, not an unusable template.
    [Test]
    public async Task NullArguments()
    {
        var store = new TemplateStore();
        var template = Assert.Throws<ArgumentNullException>(() => store.RegisterDocxTemplate<Model>((Stream) null!));
        await Assert.That(template.ParamName).IsEqualTo("template");
        Assert.Throws<ArgumentNullException>(() => store.RegisterDocxTemplate<Model>((string) null!));
        var markdown = Assert.Throws<ArgumentNullException>(() => store.RegisterMarkdownTemplate<Model>(null!));
        await Assert.That(markdown.ParamName).IsEqualTo("markdown");
        Assert.Throws<ArgumentNullException>(() => GeneratedRegistration.RegisterDocxTemplate(typeof(Model), null!));
        Assert.Throws<ArgumentNullException>(() => GeneratedRegistration.RegisterMarkdownTemplate(typeof(Model), null!));
    }

    // A generated definition is only bytes until a store first renders it, which is where it is
    // registered and so where it is rejected.
    [Test]
    public async Task GeneratedDefinitions()
    {
        GeneratedRegistration.RegisterDocxTemplate(typeof(GeneratedDocxModel), []);
        GeneratedRegistration.RegisterMarkdownTemplate(typeof(GeneratedMarkdownModel), "# {{ Title }}", Spreadsheet());

        using var output = new MemoryStream();
        var docx = await Assert.That(() => new TemplateStore().Render(new GeneratedDocxModel(), output))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(docx!.Message).Contains("The template is empty");
        await Assert.That(docx.TemplateName).IsEqualTo(nameof(GeneratedDocxModel));

        var markdown = await Assert.That(() => new TemplateStore().Render(new GeneratedMarkdownModel(), output))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(markdown!.Message).Contains("style source");
        await Assert.That(markdown.Message).Contains("spreadsheet");
        await Assert.That(output.Length).IsEqualTo(0);
    }

    [Test]
    public async Task UsableInputStillRegisters()
    {
        var store = new TemplateStore();
        using var template = new NonSeekableStream(Valid());
        store.RegisterDocxTemplate<Model>(template);
        using var docx = new MemoryStream();
        await store.Render(new Model { Title = "Hello" }, docx);
        await Assert.That(docx.Length).IsGreaterThan(0);

        using var styleSource = new NonSeekableStream(Valid());
        store.RegisterMarkdownTemplate<Model>("# {{ Title }}", styleSource);
        using var markdown = new MemoryStream();
        await store.Render(new Model { Title = "Hello" }, markdown);
        await Assert.That(markdown.Length).IsGreaterThan(0);
    }
}
