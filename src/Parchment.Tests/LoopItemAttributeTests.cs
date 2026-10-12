// ReSharper disable PartialTypeWithSinglePart
// The attribute maps are keyed on dotted paths from the model root through instance members, so a
// token that reaches an annotated member through a loop variable is an ordinary value handed to
// Fluid: it may share its paragraph and take a filter. The generator half of each case is in
// Parchment.SourceGenerator.Tests' LoopItemAttributeTests, over the same templates. An annotated
// static member is the same story, but PARCH019 keeps one out of a model compiled here.
public partial class LoopItemAttributeTests
{
    // Matched by name, the way a consumer declares it.
    [AttributeUsage(AttributeTargets.Property)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property)]
    sealed class MarkdownAttribute : Attribute;

    public class Row
    {
        public required string Item { get; init; }
    }

    public class Order
    {
        [ExcelsiorTable]
        public required List<Row> Rows { get; init; }

        public required List<string> Tags { get; init; }

        [Html]
        public required string Notes { get; init; }

        [Markdown]
        public required string Memo { get; init; }

        public List<Order> Children { get; init; } = [];
    }

    [ParchmentBindable]
    public partial class Invoice
    {
        public required List<Order> Orders { get; init; }

        [ExcelsiorTable]
        public required List<Row> Rows { get; init; }

        [Html]
        public required string Notes { get; init; }
    }

    static Order NewOrder(string name, params Order[] children) =>
        new()
        {
            Rows =
            [
                new()
                {
                    Item = name
                },
                new()
                {
                    Item = name
                }
            ],
            Tags = ["red", "green"],
            Notes = $"<b>{name}</b>",
            Memo = $"**{name}**",
            Children = [..children]
        };

    static Invoice Model() =>
        new()
        {
            Orders = [NewOrder("one", NewOrder("kid"))],
            Rows =
            [
                new()
                {
                    Item = "root"
                }
            ],
            Notes = "<b>root</b>"
        };

    const string loopItemTemplate =
        """
        {% for order in Orders %}

        n={{ order.Rows | size }} t={{ order.Tags | join: '+' }}

        {{ order.Notes | upcase }}

        {{ order.Notes | html }}

        {{ order.Memo | markdown }}

        {% for child in order.Children %}

        c={{ child.Rows | size }} {{ child.Notes | upcase }}

        {% endfor %}

        {% endfor %}
        """;

    const string rootTableInsideLoopTemplate =
        """
        {% for order in Orders %}

        n={{ Rows | size }}

        {% endfor %}
        """;

    const string rootFormatInsideLoopTemplate =
        """
        {% for order in Orders %}

        {{ Notes | upcase }}

        {% endfor %}
        """;

    static readonly string[] loopItemExpected = ["n=2 t=red+green", "<B>ONE</B>", "one", "one", "c=2 <B>KID</B>"];

    [Test]
    public async Task LoopItemMembersRenderThroughFluid() =>
        await Assert.That(await RenderDocx(loopItemTemplate, Model())).IsEquivalentTo(loopItemExpected);

    [Test]
    public async Task LoopItemMembersRenderThroughFluid_Markdown() =>
        await Assert.That(await RenderMarkdown(loopItemTemplate, Model())).IsEquivalentTo(loopItemExpected);

    // Inside a loop body a path from the root is still the mapped member, and keeps its rules.
    [Test]
    public async Task RootTableInsideLoopIsStillRejected()
    {
        using var template = DocxTemplateBuilder.Build(rootTableInsideLoopTemplate);
        var store = new TemplateStore();

        await Assert.That(() => store.RegisterDocxTemplate<Invoice>(template))
            .Throws<ParchmentRegistrationException>();
    }

    [Test]
    public async Task RootFormatInsideLoopIsStillRejected()
    {
        using var template = DocxTemplateBuilder.Build(rootFormatInsideLoopTemplate);
        var store = new TemplateStore();

        await Assert.That(() => store.RegisterDocxTemplate<Invoice>(template))
            .Throws<ParchmentRegistrationException>();
    }

    [Test]
    public async Task RootTableInsideLoopIsStillRejected_Markdown()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();

        await Assert.That(() => store.RegisterMarkdownTemplate<Invoice>(rootTableInsideLoopTemplate, styleSource))
            .Throws<ParchmentRegistrationException>();
    }

    static async Task<List<string>> RenderDocx<TModel>(string source, TModel model)
        where TModel : class
    {
        using var template = DocxTemplateBuilder.Build(source);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<TModel>(template);
        return await Render(store, model);
    }

    static async Task<List<string>> RenderMarkdown<TModel>(string source, TModel model)
        where TModel : class
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TModel>(source, styleSource);
        return await Render(store, model);
    }

    static async Task<List<string>> Render<TModel>(TemplateStore store, TModel model)
        where TModel : class
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        return body.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();
    }
}
