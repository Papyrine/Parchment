using V = DocumentFormat.OpenXml.Vml;

// A text box's paragraphs are nested inside the paragraph the box is anchored in. The generator
// reads each as its own paragraph, the same as the runtime scanner does.
public class TextBoxTokenGeneratorTests
{
    const string model =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Letter
        {
            public string Title { get; set; } = "";

            public bool Show { get; set; }

            public List<string> Tags { get; set; } = new();
        }
        """;

    // Read as part of the anchoring paragraph, the tags would share a paragraph with "Label" and
    // with each other, tripping PARCH005.
    [Test]
    public async Task BlockTagsInsideATextBox_AreNotAttributedToTheAnchoringParagraph()
    {
        var diagnostics = Run(
            new Paragraph(
                new Run(new Text("Label")),
                new Run(TextBox("{% if Show %}", "shown to {{ Title }}", "{% endif %}"))));

        await Assert.That(diagnostics).IsEmpty();
    }

    // The tags of one block share the text box as their container, so PARCH028 stays quiet.
    [Test]
    public async Task LoopInsideATextBox_SharesOneContainer()
    {
        var diagnostics = Run(
            new Paragraph(
                new Run(new Text("Label")),
                new Run(TextBox("{% for tag in Tags %}", "{{ tag }}", "{% endfor %}"))));

        await Assert.That(diagnostics).IsEmpty();
    }

    // The open tag sits in the body and the close tag in a text box anchored further down.
    [Test]
    public async Task BlockTagPairSplitByATextBox_ReportsParch028()
    {
        var diagnostics = Run(
            new Paragraph(new Run(new Text("{% if Show %}"))),
            new Paragraph(
                new Run(new Text("Label")),
                new Run(TextBox("{% endif %}"))));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH028"]);
    }

    [Test]
    public async Task TokenInsideATextBox_IsValidatedOnce()
    {
        var diagnostics = Run(
            new Paragraph(
                new Run(new Text("Label")),
                new Run(TextBox("{{ Missing }}"))));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH001"]);
    }

    const string editableModel =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Letter
        {
            [EditableField]
            public string? Reference { get; set; }
        }
        """;

    // Word stores a text box as a DrawingML shape in mc:Choice and again as VML in mc:Fallback.
    // The fallback copy of an editable token is not a second use of the member.
    [Test]
    public async Task EditableTokenInBothCopiesOfATextBox_IsOneUse()
    {
        var diagnostics = Run(editableModel, TextBoxPair("{{ Reference }}"));

        await Assert.That(diagnostics).IsEmpty();
    }

    [Test]
    public async Task EditableTokenTwiceInOneCopyOfATextBox_ReportsParch017()
    {
        var diagnostics = Run(editableModel, TextBoxPair("{{ Reference }} and {{ Reference }}"));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH017"]);
    }

    [Test]
    public async Task EditableTokenInTheBodyAndInATextBox_ReportsParch017()
    {
        var diagnostics = Run(
            editableModel,
            new Paragraph(new Run(new Text("{{ Reference }}"))),
            TextBoxPair("{{ Reference }}"));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH017"]);
    }

    [Test]
    public async Task EditableTokenWithAFilterInATextBox_ReportsParch016ForEachCopy()
    {
        var diagnostics = Run(editableModel, TextBoxPair("{{ Reference | upcase }}"));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH016", "PARCH016"]);
    }

    static Paragraph TextBoxPair(string text)
    {
        var content = $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";
        return new(
            $"""
             <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="tb"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
             """);
    }

    static List<string> Run(params Paragraph[] paragraphs) =>
        Run(model, paragraphs);

    static List<string> Run(string source, params Paragraph[] paragraphs)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Letter.parchment.docx", BuildDocx(paragraphs)));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
    }

    static Picture TextBox(params string[] paragraphs) =>
        new(
            new V.Shape(
                new V.TextBox(
                    new TextBoxContent(
                        paragraphs.Select(_ => new Paragraph(new Run(new Text(_))))))));

    static byte[] BuildDocx(Paragraph[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(paragraphs));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new(
                new RemovePersonalInformation(),
                new RemoveDateAndTime());
        }

        return stream.ToArray();
    }
}
