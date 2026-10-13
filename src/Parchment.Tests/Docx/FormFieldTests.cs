using DocumentFormat.OpenXml.Validation;

// Templates authored as Word forms carry their placeholders as legacy FORMTEXT fields rather than
// {{ tokens }}. Registration rewrites them, so such a template binds like any other docx template.
// ReSharper disable PartialTypeWithSinglePart
public partial class FormFieldTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required string Title { get; init; }
    }

    [Test]
    public async Task FormTextFieldBindsLikeAToken()
    {
        using var template = BuildForm("Title", "placeholder");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var text = await Render(store);

        await Assert.That(text).IsEqualTo("bound value");
    }

    // Word forms are usually saved as templates, so a .dotx has to bind the same as a .docx.
    [Test]
    public async Task FormTextFieldInDotxBindsLikeAToken()
    {
        using var template = BuildForm("Title", "placeholder", type: WordprocessingDocumentType.Template);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var text = await Render(store);

        await Assert.That(text).IsEqualTo("bound value");
    }

    [Test]
    public async Task FormFieldKeepsTheFormattingOfItsResult()
    {
        using var template = BuildForm("Title", "placeholder", bold: true);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var output = new MemoryStream();
        await store.Render(new Model {Title = "bound value"}, output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        var run = result.MainDocumentPart!.Document!.Body!.Descendants<Run>().Single();

        await Assert.That(run.RunProperties?.Bold).IsNotNull();
    }

    // Checkbox and dropdown fields hold values Word manages, so they are not token sites.
    [Test]
    public async Task NonTextFormFieldIsLeftAlone()
    {
        using var template = BuildForm("Title", "placeholder", textInput: false);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var text = await Render(store);

        await Assert.That(text).DoesNotContain("bound value");
        await Assert.That(text).Contains("placeholder");
    }

    // "Due date" could never be a member reference, so emitting a token would only fail later.
    [Test]
    public async Task FieldNameThatCannotBeATokenIsLeftAlone()
    {
        using var template = BuildForm("Due date", "placeholder");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var text = await Render(store);

        await Assert.That(text).DoesNotContain("bound value");
        await Assert.That(text).Contains("placeholder");
    }

    // A field result can hold another field (a PAGE or REF inside the placeholder text). The outer
    // field ends at the end that balances its begin, not at the nested field's end.
    [Test]
    public async Task FormFieldWhoseResultHoldsANestedFieldIsReplacedWhole()
    {
        var runs = new List<OpenXmlElement> {TextRun("before ")};
        runs.AddRange(FieldStart("Title"));
        runs.Add(TextRun("ph "));
        runs.AddRange(PageField());
        runs.Add(TextRun(" tail"));
        runs.Add(FieldEnd());
        runs.Add(TextRun(" after"));
        using var template = Build(new Paragraph(runs));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).IsEqualTo("before bound value after");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(0);
        await Assert.That(Validate(result)).IsEmpty();
    }

    [Test]
    public async Task FormFieldInsideAHyperlinkBindsLikeAToken()
    {
        var hyperlink = new Hyperlink(FormField("Title")) {Anchor = "target"};
        using var template = Build(new Paragraph(hyperlink));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).IsEqualTo("bound value");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(0);
        await Assert.That(body.Descendants<Hyperlink>().Single().InnerText).IsEqualTo("bound value");
        await Assert.That(Validate(result)).IsEmpty();
    }

    [Test]
    public async Task FormFieldInsideATrackedInsertionBindsLikeAToken()
    {
        var inserted = new InsertedRun(FormField("Title")) {Id = "1", Author = "author"};
        using var template = Build(new Paragraph(inserted));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).IsEqualTo("bound value");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(0);
    }

    [Test]
    public async Task FormFieldInsideAContentControlBindsLikeAToken()
    {
        var control = new SdtRun(new SdtContentRun(FormField("Title")));
        using var template = Build(new Paragraph(control));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).IsEqualTo("bound value");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(0);
    }

    // The field opens inside the hyperlink and closes after it. The token lands where the begin
    // run was and every run of the field goes, so no half of the field is left behind.
    [Test]
    public async Task FormFieldSplitAcrossAWrapperBoundaryIsReplacedWhole()
    {
        var hyperlink = new Hyperlink(FieldStart("Title")) {Anchor = "target"};
        using var template = Build(new Paragraph(hyperlink, TextRun("ph"), FieldEnd(), TextRun(" after")));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).IsEqualTo("bound value after");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(0);
        await Assert.That(Validate(result)).IsEmpty();
    }

    // A field is bounded by its paragraph. One that closes in a later paragraph is left as it is,
    // and a complete field nested in its result does not lend it an end.
    [Test]
    public async Task FormFieldClosedInALaterParagraphIsLeftAlone()
    {
        var first = new List<OpenXmlElement>(FieldStart("Title")) {TextRun("ph ")};
        first.AddRange(PageField());
        using var template = Build(new Paragraph(first), new Paragraph(TextRun("more"), FieldEnd()));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).DoesNotContain("bound value");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(6);
    }

    // A field in a tracked deletion is content the author removed, not a placeholder.
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
        using var template = Build(new Paragraph(TextRun("kept"), deleted));

        using var result = await RenderDocument(template);
        var body = result.MainDocumentPart!.Document!.Body!;

        await Assert.That(body.InnerText).DoesNotContain("bound value");
        await Assert.That(body.Descendants<FieldChar>().Count()).IsEqualTo(3);
    }

    static async Task<WordprocessingDocument> RenderDocument(Stream template)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = new MemoryStream();
        await store.Render(new Model {Title = "bound value"}, output);
        output.Position = 0;
        return WordprocessingDocument.Open(output, false);
    }

    static List<string> Validate(WordprocessingDocument document) =>
        new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();

    static Run TextRun(string text) =>
        new(new Text(text) {Space = SpaceProcessingModeValues.Preserve});

    static FieldChar Begin(string name) =>
        new FieldChar {FieldCharType = FieldCharValues.Begin}
            .AppendField(new(new FormFieldName {Val = name}, new TextInput()));

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

    static MemoryStream Build(params Paragraph[] paragraphs)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(paragraphs.Cast<OpenXmlElement>()));
        }

        stream.Position = 0;
        return stream;
    }

    static async Task<string> Render(TemplateStore store)
    {
        using var output = new MemoryStream();
        await store.Render(new Model {Title = "bound value"}, output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        return result.MainDocumentPart!.Document!.Body!.InnerText;
    }

    static MemoryStream BuildForm(string name, string result, bool bold = false, bool textInput = true, WordprocessingDocumentType type = WordprocessingDocumentType.Document)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, type))
        {
            var main = doc.AddMainDocumentPart();

            var data = new FormFieldData(new FormFieldName {Val = name});
            if (textInput)
            {
                data.AppendChild(new TextInput());
            }
            else
            {
                data.AppendChild(new CheckBox());
            }

            var resultRun = new Run(new Text(result));
            if (bold)
            {
                resultRun.RunProperties = new(new Bold());
            }

            main.Document = new(
                new Body(
                    new Paragraph(
                        new Run(
                            new FieldChar
                            {
                                FieldCharType = FieldCharValues.Begin
                            }.AppendField(data)),
                        new Run(new FieldCode(" FORMTEXT ")),
                        new Run(
                            new FieldChar
                            {
                                FieldCharType = FieldCharValues.Separate
                            }),
                        resultRun,
                        new Run(
                            new FieldChar
                            {
                                FieldCharType = FieldCharValues.End
                            }))));
        }

        stream.Position = 0;
        return stream;
    }
}

static class FieldCharExtensions
{
    public static FieldChar AppendField(this FieldChar fieldChar, FormFieldData data)
    {
        fieldChar.AppendChild(data);
        return fieldChar;
    }
}
