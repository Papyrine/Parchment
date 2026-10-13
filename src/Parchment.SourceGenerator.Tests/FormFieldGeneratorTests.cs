// The compile-time half of FormFieldTests in Parchment.Tests. Registration rewrites a legacy
// FORMTEXT field into a {{ Name }} token before anything scans the paragraph, so the generator has
// to read a paragraph as it stands after that rewrite. The reader tests here build the same field
// shapes as the runtime tests of the same name and pin the text the runtime scans for each.
public class FormFieldGeneratorTests
{
    const string source =
        """
        using Parchment;

        namespace Sample;

        public class Customer
        {
            public string Name { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Letter
        {
            public string Title { get; set; } = "";

            public bool Flag { get; set; }

            public Customer Customer { get; set; } = new();
        }
        """;

    [Test]
    public async Task FormFieldNamingAMissingMember_ReportsParch001()
    {
        var diagnostics = Run(new Paragraph(FormField("Missing")));

        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH001"]);
        await Assert.That(diagnostics.Single().GetMessage()).Contains("Missing");
    }

    [Test]
    public async Task FormFieldNamingAMember_NoDiagnostic()
    {
        var diagnostics = Run(
            new Paragraph(FormField("Title")),
            new Paragraph(FormField("Customer.Name")));

        await Assert.That(diagnostics).IsEmpty();
    }

    // The placeholder a field displays goes with the field, so a token typed there is never read.
    [Test]
    public async Task TokenInThePlaceholderOfAFormField_IsNotValidated()
    {
        var diagnostics = Run(new Paragraph([..FieldStart("Title"), TextRun("{{ Missing }}"), FieldEnd()]));

        await Assert.That(diagnostics).IsEmpty();
    }

    // Once the field is a token it is a substitution in the tag's paragraph, and the runtime
    // rejects a block tag that shares its paragraph with one, whatever the token came from.
    [Test]
    public async Task FormFieldBesideABlockTagIsATokenBesideIt()
    {
        var diagnostics = Run(
            new Paragraph([TextRun("{% if Flag %}"), ..FormField("Title")]),
            new Paragraph(TextRun("{% endif %}")));

        await Assert.That(diagnostics.Select(_ => _.Id)).Contains("PARCH005");
    }

    [Test]
    public async Task FormFieldInAHeader_IsValidated()
    {
        var bytes = GeneratorDriver.BuildDocxBytesWithHeader(["body"], []);
        using (var stream = new MemoryStream())
        {
            stream.Write(bytes, 0, bytes.Length);
            using (var doc = WordprocessingDocument.Open(stream, true))
            {
                doc.MainDocumentPart!.HeaderParts.Single().Header!.Append(new Paragraph(FormField("Missing")));
            }

            bytes = stream.ToArray();
        }

        var diagnostics = Run(bytes);

        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task FormTextFieldBindsLikeAToken()
    {
        var texts = Read(new Paragraph([TextRun("before "), ..FormField("Title"), TextRun(" after")]));

        await Assert.That(texts).IsEquivalentTo(["before {{ Title }} after"]);
    }

    [Test]
    public async Task NonTextFormFieldIsLeftAlone()
    {
        var begin = new FieldChar {FieldCharType = FieldCharValues.Begin};
        begin.AppendChild(new FormFieldData(new FormFieldName {Val = "Title"}, new CheckBox()));
        var texts = Read(
            new Paragraph(
                new Run(begin),
                new Run(new FieldCode(" FORMCHECKBOX ")),
                new Run(new FieldChar {FieldCharType = FieldCharValues.Separate}),
                TextRun("placeholder"),
                FieldEnd()));

        await Assert.That(texts).IsEquivalentTo(["placeholder"]);
    }

    [Test]
    public async Task FieldNameThatCannotBeATokenIsLeftAlone()
    {
        var texts = Read(
            new Paragraph(FormField("Due date")),
            new Paragraph(FormField("1st")),
            new Paragraph(FormField("")));

        await Assert.That(texts).IsEquivalentTo(["ph", "ph", "ph"]);
    }

    [Test]
    public async Task FormFieldWhoseResultHoldsANestedFieldIsReplacedWhole()
    {
        var texts = Read(
            new Paragraph(
            [
                TextRun("before "),
                ..FieldStart("Title"),
                TextRun("ph "),
                ..PageField(),
                TextRun(" tail"),
                FieldEnd(),
                TextRun(" after")
            ]));

        await Assert.That(texts).IsEquivalentTo(["before {{ Title }} after"]);
    }

    [Test]
    public async Task FormFieldInsideAHyperlinkBindsLikeAToken()
    {
        var texts = Read(new Paragraph(new Hyperlink(FormField("Title")) {Anchor = "target"}));

        await Assert.That(texts).IsEquivalentTo(["{{ Title }}"]);
    }

    [Test]
    public async Task FormFieldInsideATrackedInsertionBindsLikeAToken()
    {
        var texts = Read(new Paragraph(new InsertedRun(FormField("Title")) {Id = "1", Author = "author"}));

        await Assert.That(texts).IsEquivalentTo(["{{ Title }}"]);
    }

    [Test]
    public async Task FormFieldInsideAContentControlBindsLikeAToken()
    {
        var texts = Read(new Paragraph(new SdtRun(new SdtContentRun(FormField("Title")))));

        await Assert.That(texts).IsEquivalentTo(["{{ Title }}"]);
    }

    [Test]
    public async Task FormFieldSplitAcrossAWrapperBoundaryIsReplacedWhole()
    {
        var hyperlink = new Hyperlink(FieldStart("Title")) {Anchor = "target"};
        var texts = Read(new Paragraph(hyperlink, TextRun("ph"), FieldEnd(), TextRun(" after")));

        await Assert.That(texts).IsEquivalentTo(["{{ Title }} after"]);
    }

    [Test]
    public async Task FormFieldClosedInALaterParagraphIsLeftAlone()
    {
        var texts = Read(
            new Paragraph([..FieldStart("Title"), TextRun("ph "), ..PageField()]),
            new Paragraph(TextRun("more"), FieldEnd()));

        await Assert.That(texts).IsEquivalentTo(["ph 1", "more"]);
    }

    [Test]
    public async Task FormFieldInsideATrackedDeletionIsLeftAlone()
    {
        var deleted = new DeletedRun(
            new Run(Begin("Title")),
            new Run(new DeletedFieldCode(" FORMTEXT ")),
            new Run(new FieldChar {FieldCharType = FieldCharValues.Separate}),
            new Run(new DeletedText("ph")),
            FieldEnd())
        {
            Id = "1",
            Author = "author"
        };
        var texts = Read(new Paragraph(TextRun("kept"), deleted));

        await Assert.That(texts).IsEquivalentTo(["kept"]);
    }

    [Test]
    public async Task TwoFormFieldsInOneParagraphBothConvert()
    {
        var texts = Read(new Paragraph([..FormField("Title"), TextRun(" and "), ..FormField("Customer.Name")]));

        await Assert.That(texts).IsEquivalentTo(["{{ Title }} and {{ Customer.Name }}"]);
    }

    // The runtime removes the runs of a converted field, so a text box anchored in the placeholder
    // goes with them and its paragraphs are never scanned.
    [Test]
    public async Task TextBoxAnchoredInAFormFieldGoesWithTheField()
    {
        var anchor = new Run(new Picture(new Paragraph(TextRun("{{ Missing }}"))));
        var texts = Read(new Paragraph([..FieldStart("Title"), anchor, FieldEnd()]));

        await Assert.That(texts).IsEquivalentTo(["{{ Title }}"]);
    }

    static ImmutableArray<Diagnostic> Run(params OpenXmlElement[] blocks) =>
        Run(GeneratorDriver.BuildDocxBytesFromBlocks(blocks));

    static ImmutableArray<Diagnostic> Run(byte[] docx)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(source, new TemplateFile("Letter.parchment.docx", docx));
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult().Results.Single().Diagnostics;
    }

    static List<string> Read(params OpenXmlElement[] blocks)
    {
        var setup = GeneratorDriver.CreateDriverWithDocxes(
            source,
            new TemplateFile("Letter.parchment.docx", GeneratorDriver.BuildDocxBytesFromBlocks(blocks)));
        return DocxArchiveReader.Read(File.ReadAllBytes(setup.DocxPath)).Paragraphs;
    }

    static Run TextRun(string text) =>
        new(new Text(text) {Space = SpaceProcessingModeValues.Preserve});

    static FieldChar Begin(string name)
    {
        var fieldChar = new FieldChar {FieldCharType = FieldCharValues.Begin};
        fieldChar.AppendChild(new FormFieldData(new FormFieldName {Val = name}, new TextInput()));
        return fieldChar;
    }

    static Run FieldEnd() =>
        new(new FieldChar {FieldCharType = FieldCharValues.End});

    // The begin, instruction and separate runs of a FORMTEXT field; the result and end follow.
    static OpenXmlElement[] FieldStart(string name) =>
    [
        new Run(Begin(name)),
        new Run(new FieldCode(" FORMTEXT ")),
        new Run(new FieldChar {FieldCharType = FieldCharValues.Separate})
    ];

    static OpenXmlElement[] FormField(string name) =>
        [..FieldStart(name), TextRun("ph"), FieldEnd()];

    static OpenXmlElement[] PageField() =>
    [
        new Run(new FieldChar {FieldCharType = FieldCharValues.Begin}),
        new Run(new FieldCode(" PAGE ")),
        new Run(new FieldChar {FieldCharType = FieldCharValues.Separate}),
        TextRun("1"),
        FieldEnd()
    ];
}
