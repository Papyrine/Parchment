public class ListHostStyleTests
{
    static string SourcePath([CallerFilePath] string path = "") => path;

    static string ScenarioPath(string scenarioName) =>
        Path.GetFullPath(Path.Combine(
            Path.GetDirectoryName(SourcePath())!,
            "..",
            "Scenarios",
            scenarioName));

    public class Doc
    {
        public required string Title { get; init; }
        public required IReadOnlyList<string> Items { get; init; }
    }

    [Test, Explicit]
    public async Task GenerateScenarioInputDocx()
    {
        // Authors the scenario template programmatically because the host paragraph needs a
        // user-defined pStyle ("Callout") with an explicit smaller font — DocxTemplateBuilder
        // emits empty style entries and would not show a visible inheritance difference.
        var stream = BuildInputDocx();
        var path = Path.Combine(ScenarioPath("list-host-style"), "input.docx");
        await using var file = File.Create(path);
        stream.Position = 0;
        await stream.CopyToAsync(file);
    }

    [Test]
    public async Task Render()
    {
        var templatePath = Path.Combine(ScenarioPath("list-host-style"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("list-host-style-scenario", templatePath);

        var model = new Doc
        {
            Title = "Reading list",
            Items = ["Code", "Tests", "Ship"]
        };

        using var stream = new MemoryStream();
        await store.Render("list-host-style-scenario", model, stream);

        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("list-host-style"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }

    static MemoryStream BuildInputDocx()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var body = mainPart.Document.Body!;

            AddStyles(mainPart);

            body.Append(BuildPlainParagraph("{{ Title }}"));
            body.Append(BuildStyledParagraph("Callout", "{{ Items | bullet_list }}"));
            body.Append(BuildPlainParagraph("End."));

            body.Append(
                new SectionProperties(
                    new PageSize { Width = 6500, Height = 8000 },
                    new PageMargin
                    {
                        Top = 500,
                        Right = 500,
                        Bottom = 500,
                        Left = 500,
                        Header = 720,
                        Footer = 720
                    }));
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph BuildPlainParagraph(string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    static Paragraph BuildStyledParagraph(string styleId, string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }))
        {
            ParagraphProperties = new(new ParagraphStyleId { Val = styleId })
        };

    static void AddStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();

        styles.Append(BuildStyle("Normal", StyleValues.Paragraph, isDefault: true));
        styles.Append(BuildStyle("ListParagraph", StyleValues.Paragraph));

        // Callout: smaller, italic, blue. The visible font difference is what makes this
        // scenario useful — produced bullets must adopt Callout to look like the surrounding
        // text. Without the host-style inheritance, bullets fall back to ListParagraph (Normal)
        // and render in the default body size, which is jarring inside a styled block.
        var callout = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Callout"
        };
        callout.Append(new StyleName { Val = "Callout" });
        callout.Append(new BasedOn { Val = "Normal" });
        callout.Append(new StyleRunProperties(
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
            new Italic(),
            new Color { Val = "1F4E79" },
            new FontSize { Val = "18" },
            new FontSizeComplexScript { Val = "18" }));
        styles.Append(callout);

        stylesPart.Styles = styles;
    }

    static Style BuildStyle(string id, StyleValues type, bool isDefault = false)
    {
        var style = new Style
        {
            Type = type,
            StyleId = id
        };
        style.Append(new StyleName { Val = id });
        if (isDefault)
        {
            style.Default = true;
        }

        return style;
    }
}
