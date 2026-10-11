using V = DocumentFormat.OpenXml.Vml;

// A text box's paragraphs are nested inside the paragraph the box is anchored in. The generator
// reads each as its own paragraph, the same as the runtime scanner does.
public class TextBoxTokenGeneratorTests
{
    const string model =
        """
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Letter
        {
            public string Title { get; set; } = "";

            public bool Show { get; set; }
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

    [Test]
    public async Task TokenInsideATextBox_IsValidatedOnce()
    {
        var diagnostics = Run(
            new Paragraph(
                new Run(new Text("Label")),
                new Run(TextBox("{{ Missing }}"))));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH001"]);
    }

    static List<string> Run(Paragraph paragraph)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            model,
            new TemplateFile("Letter.parchment.docx", BuildDocx(paragraph)));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
    }

    static Picture TextBox(params string[] paragraphs) =>
        new(
            new V.Shape(
                new V.TextBox(
                    new TextBoxContent(
                        paragraphs.Select(_ => new Paragraph(new Run(new Text(_))))))));

    static byte[] BuildDocx(Paragraph paragraph)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(paragraph));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new(
                new RemovePersonalInformation(),
                new RemoveDateAndTime());
        }

        return stream.ToArray();
    }
}
