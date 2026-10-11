using V = DocumentFormat.OpenXml.Vml;

// A text box is anchored in a run of one paragraph and holds paragraphs of its own. Those inner
// paragraphs are token hosts in their own right, so their text must not also be read as part of
// the paragraph the box is anchored in.
// ReSharper disable PartialTypeWithSinglePart
public partial class TextBoxTokenTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required string Title { get; init; }

        public bool Show { get; init; }
    }

    [Test]
    public async Task TokenIsSubstitutedOnceWhenTheValueIsShorter()
    {
        using var template = Build(new Paragraph(new Run(VmlTextBox("{{ Title }}"))));

        var texts = await Render(template, "Al");

        await Assert.That(texts).IsEquivalentTo(["Al"]);
    }

    [Test]
    public async Task TokenIsSubstitutedOnceWhenTheValueIsLonger()
    {
        using var template = Build(new Paragraph(new Run(VmlTextBox("{{ Title }}"))));

        var texts = await Render(template, "Alexandria the Great");

        await Assert.That(texts).IsEquivalentTo(["Alexandria the Great"]);
    }

    // Word writes a DrawingML text box with a VML fallback, so the content exists twice — each
    // copy with its own paragraphs, and each substituted once.
    [Test]
    public async Task BothAlternateContentCopiesAreSubstitutedOnce()
    {
        const string xml =
            """
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="tb"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent><w:p><w:r><w:t>{{ Title }}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>{{ Title }}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
            """;
        using var template = Build(new Paragraph(xml));

        var texts = await Render(template, "Alexandria the Great");

        await Assert.That(texts).IsEquivalentTo(["Alexandria the Great", "Alexandria the Great"]);
    }

    [Test]
    public async Task AnchoringParagraphKeepsItsOwnTokens()
    {
        using var template = Build(
            new Paragraph(
                new Run(Preserved("Dear {{ Title }}")),
                new Run(VmlTextBox("Aside for {{ Title }}")),
                new Run(Preserved(", welcome"))));

        var texts = await Render(template, "Al");

        await Assert.That(texts).IsEquivalentTo(["Dear Al", "Aside for Al", ", welcome"]);
    }

    // Read as part of the anchoring paragraph, these tags would be three block tags sharing a
    // paragraph with "Label" and registration would reject the template.
    [Test]
    public async Task BlockTagsInsideATextBoxAreNotAttributedToTheAnchoringParagraph()
    {
        using var template = Build(
            new Paragraph(
                new Run(Preserved("Label")),
                new Run(VmlTextBox("{% if Show %}", "shown to {{ Title }}", "{% endif %}"))));

        var texts = await Render(template, "Al", show: true);

        await Assert.That(texts).IsEquivalentTo(["Label", "shown to Al"]);
    }

    static async Task<List<string>> Render(Stream template, string title, bool show = false)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var output = new MemoryStream();
        await store.Render(
            new Model
            {
                Title = title,
                Show = show
            },
            output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        return result.MainDocumentPart!.Document!.Body!
            .Descendants<Text>()
            .Select(_ => _.Text)
            .ToList();
    }

    static Picture VmlTextBox(params string[] paragraphs) =>
        new(
            new V.Shape(
                new V.TextBox(
                    new TextBoxContent(
                        paragraphs.Select(_ => new Paragraph(new Run(Preserved(_))))))));

    static Text Preserved(string text) =>
        new(text)
        {
            Space = SpaceProcessingModeValues.Preserve
        };

    static MemoryStream Build(Paragraph paragraph)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(paragraph));
        }

        stream.Position = 0;
        return stream;
    }
}
