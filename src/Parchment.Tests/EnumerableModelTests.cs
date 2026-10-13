// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A model that is itself a collection. Fluid would hand a template an array for it, so the render
/// wraps the model as an object: what the model declares binds, as it does for any other model.
/// Everywhere else the same type is still an array — as a member, and as a loop element.
/// </summary>
public partial class EnumerableModelTests
{
    public class Line
    {
        public required string Name { get; init; }
    }

    [ParchmentBindable]
    public partial class Order :
        List<Line>
    {
        public string Name { get; init; } = "order";
        public IEnumerable<Line> Items => this;
        public Order? Previous { get; init; }
        public List<Order> History { get; init; } = [];
    }

    [ParchmentBindable]
    public partial class Basket :
        IEnumerable<Line>
    {
        List<Line> lines = [];

        public string Name { get; init; } = "basket";

        public void Add(Line line) =>
            lines.Add(line);

        public IEnumerator<Line> GetEnumerator() =>
            lines.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() =>
            GetEnumerator();
    }

    [ParchmentBindable]
    public partial class Lookup :
        Dictionary<string, string>
    {
        public string Name { get; init; } = "lookup";
    }

    static Order BuildOrder() =>
        new()
        {
            new()
            {
                Name = "one"
            },
            new()
            {
                Name = "two"
            }
        };

    static Order BuildModel()
    {
        var order = new Order
        {
            Previous = BuildOrder(),
            History = [BuildOrder(), BuildOrder()]
        };
        order.Add(
            new()
            {
                Name = "first"
            });
        order.Add(
            new()
            {
                Name = "second"
            });
        order.Add(
            new()
            {
                Name = "third"
            });
        return order;
    }

    [Test]
    public async Task OwnMemberOfAListModelRenders()
    {
        await Assert.That(await RenderDocx("[{{ Name }}]", BuildModel())).IsEqualTo("[order]");
        await Assert.That(await RenderMarkdown("[{{ Name }}]", BuildModel())).IsEqualTo("[order]");
    }

    [Test]
    public async Task OwnMemberOfAnEnumerableModelRenders()
    {
        var basket = new Basket
        {
            new()
            {
                Name = "one"
            }
        };
        await Assert.That(await RenderDocx("[{{ Name }}]", basket)).IsEqualTo("[basket]");
        await Assert.That(await RenderMarkdown("[{{ Name }}]", basket)).IsEqualTo("[basket]");
    }

    // Fluid reads any name on a dictionary as a key, so an empty one is the case that shows it.
    [Test]
    public async Task OwnMemberOfADictionaryModelRenders()
    {
        var lookup = new Lookup
        {
            ["Name"] = "key"
        };
        await Assert.That(await RenderDocx("[{{ Name }}]", lookup)).IsEqualTo("[lookup]");
        await Assert.That(await RenderMarkdown("[{{ Name }}]", lookup)).IsEqualTo("[lookup]");
    }

    [Test]
    public async Task OwnMemberDrivesAConditional()
    {
        var docx = await RenderDocx(
            """
            {% if Name == 'order' %}

            yes

            {% endif %}
            """,
            BuildModel());
        await Assert.That(docx).IsEqualTo("yes");

        var markdown = await RenderMarkdown("{% if Name == 'order' %}yes{% endif %}", BuildModel());
        await Assert.That(markdown).IsEqualTo("yes");
    }

    // The model has no name a template can loop over, so its items are reached through a member,
    // and that member is an array like any other collection.
    [Test]
    [Arguments("{{ Items.size }}", "3")]
    [Arguments("{{ Items.first.Name }}", "first")]
    [Arguments("{{ Items.last.Name }}", "third")]
    [Arguments("{{ Previous.size }}", "2")]
    [Arguments("{{ Previous.first.Name }}", "one")]
    [Arguments("{{ Previous.last.Name }}", "two")]
    [Arguments("{{ History.size }}", "2")]
    [Arguments("{{ History.first.size }}", "2")]
    [Arguments("{{ History.first.last.Name }}", "two")]
    public async Task SameTypeAsAMemberIsAnArray(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]", BuildModel())).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]", BuildModel())).IsEqualTo($"[{expected}]");
    }

    [Test]
    public async Task ItemsOfTheModelAreLoopedThroughAMember()
    {
        var docx = await RenderDocx(
            """
            {% for line in Items %}

            {{ line.Name }}

            {% endfor %}
            """,
            BuildModel());
        await Assert.That(docx).IsEqualTo("firstsecondthird");

        var markdown = await RenderMarkdown("{% for line in Items %}{{ line.Name }}{% endfor %}", BuildModel());
        await Assert.That(markdown).IsEqualTo("firstsecondthird");
    }

    [Test]
    public async Task SameTypeAsAMemberIsALoopSource()
    {
        var docx = await RenderDocx(
            """
            {% for line in Previous %}

            {{ line.Name }}

            {% endfor %}
            """,
            BuildModel());
        await Assert.That(docx).IsEqualTo("onetwo");

        var markdown = await RenderMarkdown("{% for line in Previous %}{{ line.Name }}{% endfor %}", BuildModel());
        await Assert.That(markdown).IsEqualTo("onetwo");
    }

    [Test]
    public async Task LoopElementOfTheModelTypeIsAnArray()
    {
        var docx = await RenderDocx(
            """
            {% for order in History %}

            {{ order.size }}{{ order.first.Name }}{{ order.last.Name }}

            {% for line in order %}

            {{ line.Name }}

            {% endfor %}

            {% endfor %}
            """,
            BuildModel());
        await Assert.That(docx).IsEqualTo("2onetwoonetwo2onetwoonetwo");

        var markdown = await RenderMarkdown(
            "{% for order in History %}{{ order.size }}{{ order.first.Name }}{{ order.last.Name }}{% for line in order %}{{ line.Name }}{% endfor %}{% endfor %}",
            BuildModel());
        await Assert.That(markdown).IsEqualTo("2onetwoonetwo2onetwoonetwo");
    }

    // Past the root the type is an array, and an array has none of the members its class declares.
    [Test]
    [Arguments("{{ Previous.Name }}", "Name")]
    [Arguments("{{ Previous.Count }}", "Count")]
    [Arguments("{{ History.first.Name }}", "Name")]
    [Arguments("{% for order in History %}{{ order.Name }}{% endfor %}", "Name")]
    [Arguments("{% for order in History %}{{ order.Count }}{% endfor %}", "Count")]
    public async Task OwnMemberPastTheRootFailsRegistration(string markdown, string member)
    {
        var docxContent = markdown
            .Replace("%}{{", "%}\n\n{{")
            .Replace("}}{%", "}}\n\n{%");
        using var template = DocxTemplateBuilder.Build(docxContent);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Order>(template));
        await Assert.That(docx.Message).Contains($"'{member}'");
        await Assert.That(docx.Message).Contains("'size'");

        using var styleSource = DocxTemplateBuilder.Build();
        var fromMarkdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Order>(markdown, styleSource));
        await Assert.That(fromMarkdown.Message).Contains($"'{member}'");
        await Assert.That(fromMarkdown.Message).Contains("'size'");
    }

    // What the collection base declares has no generated accessor, and liquid's array members
    // belong to an array, which the model is not.
    [Test]
    [Arguments("Count")]
    [Arguments("Capacity")]
    [Arguments("size")]
    [Arguments("first")]
    [Arguments("last")]
    public async Task CollectionMemberOfTheModelFailsRegistration(string name)
    {
        using var template = DocxTemplateBuilder.Build($"{{{{ {name} }}}}");
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Order>(template));
        await Assert.That(docx.Message).Contains($"'{name}'");
        await Assert.That(docx.Message).Contains("the model itself");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Order>($"{{{{ {name} }}}}", styleSource));
        await Assert.That(markdown.Message).Contains($"'{name}'");
        await Assert.That(markdown.Message).Contains("the model itself");
    }

    [Test]
    public async Task CollectionMemberOfADictionaryModelIsNotALoopSource()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for key in Keys %}

            {{ key }}

            {% endfor %}
            """);
        Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Lookup>(template));

        using var styleSource = DocxTemplateBuilder.Build();
        Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Lookup>("{% for key in Keys %}{{ key }}{% endfor %}", styleSource));
        await Task.CompletedTask;
    }

    [Test]
    public async Task UnknownMemberOfTheModelKeepsThePlainMessage()
    {
        using var template = DocxTemplateBuilder.Build("{{ Nope }}");
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Order>(template));
        await Assert.That(docx.Message).Contains("'Nope' is not a member of 'Order'");
    }

    // A header is rendered through its own context in both flows.
    [Test]
    public async Task OwnMemberRendersInAHeader()
    {
        using var template = DocxTemplateBuilder.Build("body");
        AddHeader(template, "[{{ Name }}]");

        var docxStore = new TemplateStore();
        docxStore.RegisterDocxTemplate<Order>(template);
        await Assert.That(await RenderHeader(docxStore, BuildModel())).IsEqualTo("[order]");

        template.Position = 0;
        var markdownStore = new TemplateStore();
        markdownStore.RegisterMarkdownTemplate<Order>("body", template);
        await Assert.That(await RenderHeader(markdownStore, BuildModel())).IsEqualTo("[order]");
    }

    static void AddHeader(MemoryStream template, string text)
    {
        template.Position = 0;
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader1");
            headerPart.Header = new(new Paragraph(new Run(new Text(text))));
            var body = mainPart.Document!.Body!;
            var section = body.Elements<SectionProperties>().LastOrDefault();
            if (section == null)
            {
                section = new();
                body.AppendChild(section);
            }

            section.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdHeader1"
                });
        }

        template.Position = 0;
    }

    static async Task<string> RenderHeader<T>(TemplateStore store, T model)
        where T : class
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.HeaderParts.Single().Header!.InnerText;
    }

    static async Task<string> RenderDocx<T>(string content, T model)
        where T : class
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);
        return await RenderText(store, model);
    }

    static async Task<string> RenderMarkdown<T>(string markdown, T model)
        where T : class
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<T>(markdown, styleSource);
        return await RenderText(store, model);
    }

    static async Task<string> RenderText<T>(TemplateStore store, T model)
        where T : class
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
