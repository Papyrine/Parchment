// Templates authored as Word forms carry their placeholders as legacy FORMTEXT fields rather than
// {{ tokens }}. Registration rewrites them, so such a template binds like any other docx template.
public class FormFieldTests
{
    public class Model
    {
        public required string Title { get; init; }
    }

    [Test]
    public async Task FormTextFieldBindsLikeAToken()
    {
        using var template = BuildForm("Title", "placeholder");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>("form", template);

        var text = await Render(store);

        await Assert.That(text).IsEqualTo("bound value");
    }

    [Test]
    public async Task FormFieldKeepsTheFormattingOfItsResult()
    {
        using var template = BuildForm("Title", "placeholder", bold: true);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>("form", template);

        using var output = new MemoryStream();
        await store.Render("form", new Model {Title = "bound value"}, output);
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
        store.RegisterDocxTemplate<Model>("form", template);

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
        store.RegisterDocxTemplate<Model>("form", template);

        var text = await Render(store);

        await Assert.That(text).DoesNotContain("bound value");
        await Assert.That(text).Contains("placeholder");
    }

    static async Task<string> Render(TemplateStore store)
    {
        using var output = new MemoryStream();
        await store.Render("form", new Model {Title = "bound value"}, output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        return result.MainDocumentPart!.Document!.Body!.InnerText;
    }

    static MemoryStream BuildForm(string name, string result, bool bold = false, bool textInput = true)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
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
