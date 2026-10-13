// A token inside a w:smartTag is the paragraph's token. The generator reads the part as plain xml
// and so sees through the smart tag unaided; the runtime scanner has to retype it first. These pin
// that the two agree.
public class SmartTagTokenGeneratorTests
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

    const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string smartTagOpen = """<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">""";
    const string smartTagClose = "</w:smartTag>";

    [Test]
    public async Task UnknownMemberInsideASmartTag_ReportsParch001()
    {
        var diagnostics = Run(P(smartTagOpen + smartTagOpen + R("{{ Missing }}") + smartTagClose + smartTagClose));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task UnknownMemberSplitAcrossASmartTag_ReportsParch001()
    {
        var diagnostics = Run(P(R("{{ Mis") + smartTagOpen + R("si") + smartTagClose + R("ng }}")));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task BlockTagsInsideSmartTags_AreSeen()
    {
        var diagnostics = Run(
            P(smartTagOpen + R("{% if Show %}") + smartTagClose),
            P(R("shown to ") + smartTagOpen + R("{{ Title }}") + smartTagClose),
            P(R("{% end") + smartTagOpen + R("if") + smartTagClose + R(" %}")));

        await Assert.That(diagnostics).IsEmpty();
    }

    // Unseen, the open tag would leave no block for the close tag to end.
    [Test]
    public async Task BlockTagInsideASmartTag_OpensTheBlockItsCloseTagEnds()
    {
        var diagnostics = Run(
            P(smartTagOpen + R("{% if Show %}") + smartTagClose),
            P(R("{{ Missing }}")),
            P(R("{% endif %}")));

        await Assert.That(diagnostics).IsEquivalentTo(["PARCH001"]);
    }

    static string P(string content) =>
        $"""<w:p xmlns:w="{wordNamespace}">{content}</w:p>""";

    static string R(string text) =>
        $"""<w:r><w:t xml:space="preserve">{text}</w:t></w:r>""";

    static List<string> Run(params string[] paragraphs)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            model,
            new TemplateFile("Letter.parchment.docx", BuildDocx(paragraphs)));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
        return result.Results.Single().Diagnostics.Select(_ => _.Id).ToList();
    }

    static byte[] BuildDocx(string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(paragraphs.Select(_ => new Paragraph(_))));

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new(
                new RemovePersonalInformation(),
                new RemoveDateAndTime());
        }

        return stream.ToArray();
    }
}
