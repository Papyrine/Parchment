// ReSharper disable PartialTypeWithSinglePart
// ReSharper disable UnusedMember.Global

/// <summary>
/// An enum reaches a template as the text it prints as, so nothing can be read off it: its
/// constants are public static fields to reflection, and a path through one registered and then
/// rendered blank.
/// </summary>
public partial class EnumMemberTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public Status Status { get; init; } = Status.Final;
        public Status? Maybe { get; init; } = Status.Final;
        public Access Access { get; init; } = Access.Write;
        public Phase Stage { get; init; } = Phase.Done;
        public DayOfWeek Day { get; init; } = DayOfWeek.Monday;
        public List<Status> History { get; init; } = [Status.Draft, Status.Final];
        public Holder Holder { get; init; } = new();

        public enum Phase
        {
            Open,
            Done
        }
    }

    public class Holder
    {
        public Status Status { get; init; } = Status.Final;
    }

    public enum Status
    {
        Draft,
        Final
    }

    [Flags]
    public enum Access
    {
        Read = 1,
        Write = 2
    }

    [Test]
    [Arguments("{{ Status }}", "Final")]
    [Arguments("{{ Maybe }}", "Final")]
    [Arguments("{{ Access }}", "Write")]
    [Arguments("{{ Stage }}", "Done")]
    [Arguments("{{ Day }}", "Monday")]
    [Arguments("{{ Holder.Status }}", "Final")]
    public async Task EnumRegistersAndRenders(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    // What the registration message points to instead of a constant.
    [Test]
    public async Task ComparingAgainstTheTextWorks()
    {
        var docx = await RenderDocx(
            """
            {% if Status == "Final" %}

            final

            {% else %}

            not final

            {% endif %}
            """);
        await Assert.That(docx).IsEqualTo("final");

        var markdown = await RenderMarkdown("""{% if Status == "Final" %}final{% else %}not final{% endif %}""");
        await Assert.That(markdown).IsEqualTo("final");
    }

    [Test]
    public async Task ComparingALoopVariableAgainstTheTextWorks()
    {
        var docx = await RenderDocx(
            """
            {% for entry in History %}

            {% if entry == "Final" %}

            final

            {% endif %}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("final");

        var markdown = await RenderMarkdown("""{% for entry in History %}{% if entry == "Final" %}final{% endif %}{% endfor %}""");
        await Assert.That(markdown).IsEqualTo("final");
    }

    // The first test pins that Fluid binds none of these, which is the only reason the second may
    // reject them. `value__` is the field every enum stores its number in, public to reflection.
    [Test]
    [Arguments("Status.Final")]
    [Arguments("Status.Draft")]
    [Arguments("Status.value__")]
    [Arguments("Maybe.Final")]
    [Arguments("Maybe.Value")]
    [Arguments("Maybe.Value.Final")]
    [Arguments("Access.Write")]
    [Arguments("Stage.Done")]
    [Arguments("Day.Monday")]
    [Arguments("Holder.Status.Final")]
    public async Task ConstantRendersBlankInFluid(string path)
    {
        SharedFluid.Parser.TryParse($"[{{{{ {path} }}}}]", out var template, out _);
        var context = new TemplateContext(new Doc(), SharedFluid.Options, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(context)).IsEqualTo("[]");
    }

    [Test]
    [Arguments("Status.Final", "Status")]
    [Arguments("Status.Draft", "Status")]
    [Arguments("Status.value__", "Status")]
    [Arguments("Maybe.Final", "Status")]
    [Arguments("Maybe.Value", "Status")]
    [Arguments("Access.Write", "Access")]
    [Arguments("Stage.Done", "Phase")]
    [Arguments("Day.Monday", "DayOfWeek")]
    [Arguments("Holder.Status.Final", "Status")]
    public async Task ConstantFailsRegistrationNamingTheEnum(string path, string enumName)
    {
        var member = path[(path.LastIndexOf('.') + 1)..];

        using var template = DocxTemplateBuilder.Build($"{{{{ {path} }}}}");
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains($"Member '{member}' of '{enumName}'");
        await Assert.That(docx.Message).Contains("enum");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>($"{{{{ {path} }}}}", styleSource));
        await Assert.That(markdown.Message).Contains($"Member '{member}' of '{enumName}'");
        await Assert.That(markdown.Message).Contains("enum");
    }

    // The reported template: with Status = Final this registered and printed "not final".
    [Test]
    public async Task ConstantInAConditionFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Status == Status.Final %}

            final

            {% else %}

            not final

            {% endif %}
            """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("Member 'Final' of 'Status'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>("{% if Status == Status.Final %}final{% else %}not final{% endif %}", styleSource));
        await Assert.That(markdown.Message).Contains("Member 'Final' of 'Status'");
    }

    [Test]
    public async Task ConstantThroughALoopVariableFailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for entry in History %}

            {{ entry.Final }}

            {% endfor %}
            """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains("Member 'Final' of 'Status'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>("{% for entry in History %}{{ entry.Final }}{% endfor %}", styleSource));
        await Assert.That(markdown.Message).Contains("Member 'Final' of 'Status'");
    }

    [Test]
    public async Task ConstantInAHeaderFailsRegistration()
    {
        using var template = new MemoryStream();
        using (var document = WordprocessingDocument.Create(template, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
            headerPart.Header = new(new Paragraph(new Run(new Text("{{ Status.Final }}"))));
            mainPart.Document = new(
                new Body(
                    new Paragraph(new Run(new Text("{{ Status }}"))),
                    new SectionProperties(
                        new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdHeader"
                        })));
        }

        template.Position = 0;
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(exception.Message).Contains("Member 'Final' of 'Status'");
    }

    static async Task<string> RenderDocx(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return await RenderText(store);
    }

    static async Task<string> RenderMarkdown(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return await RenderText(store);
    }

    static async Task<string> RenderText(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(new Doc(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
