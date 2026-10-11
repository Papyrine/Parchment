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

    static List<string> Run(params Paragraph[] paragraphs)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            model,
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
