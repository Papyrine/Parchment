// ReSharper disable PartialTypeWithSinglePart
// The generator emits accessors for every type a model reaches, but the per-model maps — formats,
// string lists, tables, editable fields — only for the model itself. A type it saw only as a
// member has the first and none of the second, so it is not a model: used as one it would bind,
// and render with every attribute ignored.
public partial class MemberTypeAsModelTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    // Deliberately NOT [ParchmentBindable]: reachable only through Holder.
    public class Customer
    {
        [Html]
        public string Bio { get; set; } = "<b>bold</b>";

        public List<string> Tags { get; set; } = ["a", "b"];

        [EditableField]
        public string Note { get; set; } = "n";
    }

    [ParchmentBindable]
    public partial class Holder
    {
        public Customer Customer { get; set; } = new();

        public List<Customer> Customers { get; set; } = [new()];
    }

    // Derives from a model without being one: the attribute is not inherited, and the generator
    // emits nothing for it.
    public class DerivedHolder : Holder;

    const string notGenerated = "has no pre-compiled Parchment accessors";

    [Test]
    public async Task DocxRegistrationRejectsAMemberType()
    {
        using var template = DocxTemplateBuilder.Build("{{ Bio }}\n\n{{ Tags }}\n\nNote: {{ Note }}");
        var store = new TemplateStore();

        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Customer>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains(notGenerated);
        await Assert.That(exception.Message).Contains(typeof(Customer).FullName!);
    }

    [Test]
    public async Task MarkdownRegistrationRejectsAMemberType()
    {
        var store = new TemplateStore();

        var exception = await Assert.That(
                () => store.RegisterMarkdownTemplate<Customer>("{{ Bio }}\n\n{{ Tags }}\n"))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains(notGenerated);
    }

    [Test]
    public async Task ExtractionRejectsAMemberType()
    {
        using var stream = DocxTemplateBuilder.Build("Note: n");

        var exception = await Assert.That(
                () => ParchmentExtractor.Extract<Customer>(stream))
            .Throws<ParchmentExtractionException>();
        await Assert.That(exception!.Message).Contains(notGenerated);
    }

    [Test]
    public async Task RegistrationRejectsATypeDerivedFromAModel()
    {
        var store = new TemplateStore();

        var exception = await Assert.That(
                () => store.RegisterMarkdownTemplate<DerivedHolder>("{{ Customer.Note }}"))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains(notGenerated);
    }

    [Test]
    public async Task RenderOfAMemberTypeFindsNoTemplate()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Holder>("{{ Customer.Note }}");

        using var output = new MemoryStream();
        await Assert.That(
                () => store.Render(new Customer(), output))
            .Throws<ParchmentRenderException>();
    }

    [Test]
    public async Task MemberTypeStillBindsAsAMemberAndALoopElement()
    {
        using var template = DocxTemplateBuilder.Build(
            "{{ Customer.Bio }}\n\n{{ Customer.Tags }}\n\nNote: {{ Customer.Note }}\n\n{% for c in Customers %}\n\nEach: {{ c.Note }}\n\n{% endfor %}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Holder>(template);

        using var output = new MemoryStream();
        await store.Render(new Holder(), output);
        var text = BodyText(output);

        // [Html] applied: the markup is gone and the word is left.
        await Assert.That(text).Contains("bold");
        await Assert.That(text).DoesNotContain("<b>");
        await Assert.That(text).Contains("Note: n");
        await Assert.That(text).Contains("Each: n");

        store.RegisterMarkdownTemplate<Holder>("{{ Customer.Bio }}\n\n{% for c in Customers %}Each: {{ c.Note }}{% endfor %}\n");
        using var markdown = new MemoryStream();
        await store.Render(new Holder(), markdown);
        text = BodyText(markdown);

        await Assert.That(text).Contains("bold");
        await Assert.That(text).DoesNotContain("<b>");
        await Assert.That(text).Contains("Each: n");
    }

    static string BodyText(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        return string.Join('\n', doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().Select(_ => _.InnerText));
    }
}
