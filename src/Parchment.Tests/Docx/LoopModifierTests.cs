// ReSharper disable PartialTypeWithSinglePart
public partial class LoopModifierTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required IReadOnlyList<string> Items { get; init; }
        public int Max { get; init; }
        public int Skip { get; init; }
    }

    static readonly string[] five = ["a", "b", "c", "d", "e"];

    [Test]
    [Arguments("limit:2", "a,b")]
    [Arguments("limit:0", "")]
    [Arguments("limit:9", "a,b,c,d,e")]
    // A negative limit drops that many from the end, as it does in liquid's own loop.
    [Arguments("limit:-2", "a,b,c")]
    [Arguments("offset:3", "d,e")]
    [Arguments("offset:9", "")]
    [Arguments("reversed", "e,d,c,b,a")]
    [Arguments("offset:1 limit:2", "b,c")]
    [Arguments("limit:2 offset:1", "b,c")]
    // The window is taken first and then reversed, whatever order the modifiers are written in.
    [Arguments("reversed limit:2", "b,a")]
    [Arguments("limit:2 offset:1 reversed", "c,b")]
    [Arguments("limit: Max", "a,b,c")]
    [Arguments("offset: Skip", "c,d,e")]
    [Arguments("offset: Skip limit: Max", "c,d,e")]
    // Liquid's own members resolve in a modifier as they do in a token.
    [Arguments("limit: Items.size", "a,b,c,d,e")]
    [Arguments("offset: Items.size", "")]
    public async Task MatchesTheMarkdownFlow(string modifiers, string expected)
    {
        var model = new Model
        {
            Items = five,
            Max = 3,
            Skip = 2
        };

        var docx = await RenderDocx($"{{% for item in Items {modifiers} %}}", model);
        await Assert.That(docx).IsEqualTo(expected);

        // The markdown flow hands the same tag to liquid, so it is the reference for what the
        // modifiers mean.
        var markdown = await RenderMarkdown($"{{% for item in Items {modifiers} %}}", model);
        await Assert.That(markdown).IsEqualTo(expected);
    }

    [Test]
    public async Task NestedLoopsWindowIndependently()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for group in Groups limit:1 %}

            {% for item in group.Items reversed %}

            {{ group.Name }}-{{ item }}

            {% endfor %}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<LoopTests.NestedModel>(template);
        using var stream = new MemoryStream();
        await store.Render(
            new LoopTests.NestedModel
            {
                Groups =
                [
                    new()
                    {
                        Name = "one",
                        Items = ["a", "b"]
                    },
                    new()
                    {
                        Name = "two",
                        Items = ["c"]
                    }
                ]
            },
            stream);

        await Assert.That(Paragraphs(stream)).IsEqualTo("one-b,one-a");
    }

    [Test]
    public async Task UnknownMemberInLimit_FailsRegistration()
    {
        using var template = Template("{% for item in Items limit: Missing %}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("Missing");
    }

    [Test]
    public async Task UnknownMemberInOffset_FailsRegistration()
    {
        using var template = Template("{% for item in Items offset: Missing %}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("Missing");
    }

    // A modifier is evaluated before the loop variable exists, so it cannot refer to it.
    [Test]
    public async Task LoopVariableInLimit_FailsRegistration()
    {
        using var template = Template("{% for item in Items limit: item %}");
        var store = new TemplateStore();
        await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
    }

    // `offset: continue` resumes where the previous loop over the same source stopped, which liquid
    // tracks in its own loop scope. The docx flow has no such scope, so it says so instead of
    // reading `continue` as a model member or starting again from the first item.
    [Test]
    public async Task OffsetContinue_FailsRegistrationWithAnExplanation()
    {
        using var template = Template("{% for item in Items offset: continue %}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Model>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("offset: continue");
        await Assert.That(exception.Message).DoesNotContain("is not a member of");
    }

    static MemoryStream Template(string openTag) =>
        DocxTemplateBuilder.Build(
            $$$"""
               {{{openTag}}}

               {{ item }}

               {% endfor %}
               """);

    static async Task<string> RenderDocx(string openTag, Model model)
    {
        using var template = Template(openTag);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        return Paragraphs(stream);
    }

    static async Task<string> RenderMarkdown(string openTag, Model model)
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>($"{openTag}\n{{{{ item }}}}\n\n{{% endfor %}}");
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        return Paragraphs(stream);
    }

    static string Paragraphs(MemoryStream stream)
    {
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0);
        return string.Join(',', paragraphs);
    }
}
