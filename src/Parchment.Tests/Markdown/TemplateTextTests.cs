// ReSharper disable PartialTypeWithSinglePart

using DocumentFormat.OpenXml.Validation;

/// <summary>
/// Text written as markdown reaches the document as the text a reader expects — an entity as its
/// character, a bracket that never closed as a bracket, angle quotes as angle quotes — in a markdown
/// template and wherever the docx flow renders a markdown value.
/// </summary>
public partial class TemplateTextTests
{
    // Matched by name, the way a consumer declares it — see readme "Html and Markdown properties".
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class TextModel
    {
        public string Plain { get; init; } = "";

        [Markdown]
        public string Marked { get; init; } = "";

        public TokenValue Token { get; init; } = "";
    }

    const string source = "AT&amp;T &copy; 2026, see [note and <<quoted>>";
    const string rendered = "AT&T \u00A9 2026, see [note and \u00ABquoted\u00BB";

    const string blocks =
        """
        # Terms &amp; Conditions

        AT&amp;T &copy; 2026, see [note and <<quoted>>

        - P&amp;L [open
        - **Q&amp;A** &#8212; <<both>>

        | Name | Note |
        | --- | --- |
        | R&amp;D | a &lt; b [c |
        """;

    static readonly string[] blocksRendered =
    [
        "Terms & Conditions",
        rendered,
        "P&L [open",
        "Q&A \u2014 \u00ABboth\u00BB",
        "Name",
        "Note",
        "R&D",
        "a < b [c"
    ];

    [Test]
    public async Task MarkdownTemplate()
    {
        var paragraphs = await RenderMarkdown(blocks, new());

        await Assert.That(paragraphs).IsEquivalentTo(blocksRendered);
    }

    // The three ways a value says it is markdown, each rendered by the same walk as the template.
    [Test]
    [Arguments("{{ Marked }}")]
    [Arguments("{{ Plain | markdown }}")]
    [Arguments("{{ Token }}")]
    public async Task MarkdownValueInAMarkdownTemplate(string token)
    {
        var paragraphs = await RenderMarkdown(token, Model(source));

        await Assert.That(paragraphs).IsEquivalentTo([rendered]);
    }

    [Test]
    [Arguments("{{ Marked }}")]
    [Arguments("{{ Plain | markdown }}")]
    [Arguments("{{ Token }}")]
    public async Task MarkdownValueInADocxTemplate(string token)
    {
        var paragraphs = await RenderDocx(token, Model(source));

        await Assert.That(paragraphs).IsEquivalentTo([rendered]);
    }

    [Test]
    [Arguments("{{ Marked }}")]
    [Arguments("{{ Plain | markdown }}")]
    [Arguments("{{ Token }}")]
    public async Task MarkdownBlocksInADocxTemplate(string token)
    {
        var paragraphs = await RenderDocx(token, Model(blocks));

        await Assert.That(paragraphs).IsEquivalentTo(blocksRendered);
    }

    // An inline value is spliced into the paragraph its token sits in rather than replacing it.
    [Test]
    public async Task MarkdownValueSplicedIntoADocxParagraph()
    {
        var paragraphs = await RenderDocx("Before {{ Marked }} after", Model(source));

        await Assert.That(paragraphs).IsEquivalentTo([$"Before {rendered} after"]);
    }

    static TextModel Model(string markdown) =>
        new()
        {
            Plain = markdown,
            Marked = markdown,
            Token = new MarkdownToken(markdown)
        };

    static async Task<List<string>> RenderMarkdown(string template, TextModel model)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TextModel>(template, styleSource);
        return await Paragraphs(store, model);
    }

    static async Task<List<string>> RenderDocx(string template, TextModel model)
    {
        using var templateStream = DocxTemplateBuilder.Build(template);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<TextModel>(templateStream);
        return await Paragraphs(store, model);
    }

    static async Task<List<string>> Paragraphs(TemplateStore store, TextModel model)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}");
        await Assert.That(errors).IsEmpty();

        return document.MainDocumentPart!.Document!.Body!
            .Descendants<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();
    }
}
