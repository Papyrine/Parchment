using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

// Word stores a text box twice: a DrawingML shape in mc:Choice, which is the copy it shows and
// edits, and a VML copy of the same content in mc:Fallback. An editable token typed once into the
// box is therefore in the template twice. The pair is one use of the member: both copies render a
// control, and extraction reads the choice.
// ReSharper disable PartialTypeWithSinglePart
public partial class EditableTextBoxTests
{
    public class Line
    {
        [EditableField]
        public string? Name { get; set; }
    }

    [ParchmentBindable]
    public partial class Model
    {
        [EditableField]
        public string? Reference { get; set; }

        [EditableField]
        public bool Approved { get; set; }

        [EditableField]
        public List<Line> Lines { get; set; } = [];
    }

    [Test]
    public async Task TokenInATextBoxRendersAControlInBothCopies()
    {
        using var template = Build(TextBoxPair("Ref: {{ Reference }}"));

        using var stream = await Render(template, new() { Reference = "PO-1" });

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var controls = Controls(body, "Reference");
        await Assert.That(controls.Count).IsEqualTo(2);
        await Assert.That(controls.Count(InFallback)).IsEqualTo(1);
        await Assert.That(controls.Select(_ => _.InnerText)).IsEquivalentTo(["PO-1", "PO-1"]);

        // Each copy is its own editable range with its own ids.
        var sdtIds = controls.Select(_ => _.SdtProperties!.GetFirstChild<SdtId>()!.Val!.Value).ToList();
        await Assert.That(sdtIds.Distinct().Count()).IsEqualTo(2);
        var permIds = body.Descendants<PermStart>().Select(_ => _.Id!.Value).ToList();
        await Assert.That(permIds.Count).IsEqualTo(2);
        await Assert.That(permIds.Distinct().Count()).IsEqualTo(2);

        await Assert.That(Errors(doc)).IsEmpty();
    }

    [Test]
    public async Task ExtractionReadsTheChoiceCopy()
    {
        using var template = Build(TextBoxPair("{{ Reference }} {{ Approved }}"));
        using var stream = await Render(template, new() { Reference = "PO-1" });

        // What Word leaves when the user edits the box and the fallback is not rewritten to match.
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            foreach (var control in Controls(body, "Reference"))
            {
                var text = control.Descendants<Text>().Single();
                if (InFallback(control))
                {
                    text.Text = "stale";
                }
                else
                {
                    text.Text = "PO-2";
                }
            }

            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<Model>(stream);

        await Assert.That(result.Fields.Any(_ => _.State == FieldState.Duplicate)).IsFalse();
        await Assert.That(result.Fields.Count(_ => _.Path == "Reference")).IsEqualTo(1);
        var model = new Model();
        result.ApplyTo(model);
        await Assert.That(model.Reference).IsEqualTo("PO-2");
        await Assert.That(model.Approved).IsFalse();
    }

    // A text box saved by an old version of Word, or built by hand, is VML with no DrawingML
    // twin. There is no fallback to skip, so its control is the one read.
    [Test]
    public async Task VmlOnlyTextBoxRoundTrips()
    {
        using var template = Build(
            new Paragraph(
                new Run(
                    new Picture(
                        new V.Shape(
                            new V.TextBox(
                                new TextBoxContent(
                                    new Paragraph(new Run(new Text("{{ Reference }}"))))))))));

        using var stream = await Render(template, new() { Reference = "PO-1" });
        var result = ParchmentExtractor.Extract<Model>(stream);

        var model = new Model();
        result.ApplyTo(model);
        await Assert.That(model.Reference).IsEqualTo("PO-1");
    }

    [Test]
    public async Task TokenTwiceInOneCopyIsStillRejected()
    {
        using var template = Build(TextBoxPair("{{ Reference }} and {{ Reference }}"));
        var store = new TemplateStore();

        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }

    [Test]
    public async Task TokenInTheBodyAndInATextBoxIsStillRejected()
    {
        using var template = Build(
            new Paragraph(new Run(new Text("{{ Reference }}"))),
            TextBoxPair("{{ Reference }}"));
        var store = new TemplateStore();

        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("more than once");
    }

    [Test]
    public async Task FilterInATextBoxIsRejectedForWhatItIs()
    {
        using var template = Build(TextBoxPair("{{ Reference | upcase }}"));
        var store = new TemplateStore();

        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("plain member-access");
    }

    // The whole loop sits inside the box, so each copy holds a repeating section of its own.
    [Test]
    public async Task CollectionLoopInATextBoxRoundTrips()
    {
        using var template = Build(TextBoxPair("{% for line in Lines %}", "{{ line.Name }}", "{% endfor %}"));
        var source = new Model
        {
            Lines =
            [
                new() { Name = "first" },
                new() { Name = "second" }
            ]
        };

        using var stream = await Render(template, source);

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sections = body.Descendants<SdtElement>()
                .Where(_ => _.SdtProperties?.GetFirstChild<W15.SdtRepeatedSection>() != null)
                .ToList();
            await Assert.That(sections.Count).IsEqualTo(2);
            await Assert.That(Errors(doc)).IsEmpty();

            foreach (var text in sections.Single(InFallback).Descendants<Text>())
            {
                text.Text = "stale";
            }

            doc.Save();
        }

        stream.Position = 0;
        var result = ParchmentExtractor.Extract<Model>(stream);
        var model = new Model();
        result.ApplyTo(model);

        await Assert.That(result.Fields.Any(_ => _.State == FieldState.Duplicate)).IsFalse();
        await Assert.That(model.Lines.Select(_ => _.Name!)).IsEquivalentTo(["first", "second"]);
    }

    static List<SdtElement> Controls(Body body, string tag) =>
        body.Descendants<SdtElement>()
            .Where(_ => _.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag)
            .ToList();

    static bool InFallback(OpenXmlElement element) =>
        element.Ancestors<AlternateContentFallback>().Any();

    static List<string> Errors(WordprocessingDocument doc) =>
        new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();

    static async Task<MemoryStream> Render(Stream template, Model model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = new MemoryStream();
        await store.Render(model, output);
        output.Position = 0;
        return output;
    }

    // The shape Word writes for an inline text box, with the same paragraphs in both copies.
    static Paragraph TextBoxPair(params string[] paragraphs)
    {
        var content = string.Concat(paragraphs.Select(_ => $"<w:p><w:r><w:t xml:space=\"preserve\">{_}</w:t></w:r></w:p>"));
        return new(
            $"""
             <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2743200" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="Text Box 1" style="width:216pt;height:72pt"><v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
             """);
    }

    static MemoryStream Build(params Paragraph[] paragraphs)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(paragraphs));
        }

        stream.Position = 0;
        return stream;
    }
}
