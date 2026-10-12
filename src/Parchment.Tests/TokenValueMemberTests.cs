// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A token value is a thing to print, not an object to read, so a path that continues into one is
/// an unknown member to every validator, whichever token type the member is declared as. The
/// generator's half is <c>TokenValueMemberGeneratorTests</c>, which runs the same paths.
/// </summary>
public partial class TokenValueMemberTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public HtmlToken Html { get; init; } = new("<b>html</b>");
        public MarkdownToken Markdown { get; init; } = new("**markdown**");
        public OpenXmlToken OpenXml { get; init; } = new(_ => []);
        public MutateToken Mutate { get; init; } = new((_, _) => { });
        public LabelledToken Labelled { get; init; } = new();
        public TokenValue Any { get; init; } = "any";
        public Section Nested { get; init; } = new();
        public List<Section> Sections { get; init; } = [new()];
        public List<HtmlToken> Tokens { get; init; } = [new("<b>listed</b>")];
    }

    public class Section
    {
        public HtmlToken Html { get; init; } = new("<b>html</b>");
        public MarkdownToken Markdown { get; init; } = new("**markdown**");
        public OpenXmlToken OpenXml { get; init; } = new(_ => []);
        public MutateToken Mutate { get; init; } = new((_, _) => { });
        public LabelledToken Labelled { get; init; } = new();
        public TokenValue Any { get; init; } = "any";
    }

    // A token type of the model's own is a token all the same, and what it declares is no more
    // reachable than what it inherits.
    public class LabelledToken() :
        HtmlToken("<b>labelled</b>")
    {
        public string Label { get; init; } = "label";
    }

    // Every public member the token types have, the one a derived token adds, the property of the
    // internal text token, and liquid's own `size`, which a string would answer.
    [Test]
    [Arguments("Html.Source")]
    [Arguments("Markdown.Source")]
    [Arguments("OpenXml.Render")]
    [Arguments("Mutate.Apply")]
    [Arguments("Labelled.Source")]
    [Arguments("Labelled.Label")]
    [Arguments("Any.Source")]
    [Arguments("Any.Value")]
    [Arguments("Html.size")]
    public async Task PathIntoATokenFailsRegistration(string path)
    {
        var member = path[(path.IndexOf('.') + 1)..];

        await AssertRejected($"{{{{ {path} }}}}", $"{{{{ {path} }}}}", member);
        await AssertRejected($"{{{{ Nested.{path} }}}}", $"{{{{ Nested.{path} }}}}", member);
        await AssertRejected(
            $"{{% for section in Sections %}}\n\n{{{{ section.{path} }}}}\n\n{{% endfor %}}",
            $"{{% for section in Sections %}}{{{{ section.{path} }}}}{{% endfor %}}",
            member);
    }

    [Test]
    public async Task PathIntoALoopVariableThatIsATokenFailsRegistration() =>
        await AssertRejected(
            "{% for token in Tokens %}\n\n{{ token.Source }}\n\n{% endfor %}",
            "{% for token in Tokens %}{{ token.Source }}{% endfor %}",
            "Source");

    [Test]
    public async Task PathIntoATokenInAConditionFailsRegistration() =>
        await AssertRejected(
            "{% if Html.Source %}\n\nsome\n\n{% endif %}",
            "{% if Html.Source %}some{% endif %}",
            "Source");

    // The reason the paths above may be rejected: the markdown flow hands Fluid the text a token
    // renders as, and that text has no such member. Only a markdown token can be shown here: the
    // others are parked against a marker, which needs a render in progress.
    [Test]
    [Arguments("Markdown.Source")]
    [Arguments("Nested.Markdown.Source")]
    public async Task MemberOfATokenRendersBlankInMarkdownFluid(string path)
    {
        SharedFluid.Parser.TryParse($"[{{{{ {path} }}}}]", out var template, out _);
        var context = new TemplateContext(new Doc(), SharedFluid.MarkdownOptions, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(context)).IsEqualTo("[]");
    }

    // No accessor is generated for a token type either, so nothing reads one through liquid.
    [Test]
    [Arguments("Html.Source")]
    [Arguments("Markdown.Source")]
    [Arguments("OpenXml.Render")]
    [Arguments("Mutate.Apply")]
    [Arguments("Labelled.Source")]
    [Arguments("Labelled.Label")]
    public async Task MemberOfATokenRendersBlankInDocxFluid(string path)
    {
        SharedFluid.Parser.TryParse($"[{{{{ {path} }}}}]", out var template, out _);
        var context = new TemplateContext(new Doc(), SharedFluid.Options, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(context)).IsEqualTo("[]");
    }

    // The token itself is untouched, wherever it is reached from.
    [Test]
    [Arguments("Html", "html")]
    [Arguments("Markdown", "markdown")]
    [Arguments("Labelled", "labelled")]
    public async Task TokenStillRenders(string path, string expected)
    {
        await Assert.That(await RenderDocx($"{{{{ {path} }}}}")).IsEqualTo(expected);
        await Assert.That(await RenderMarkdown($"{{{{ {path} }}}}")).IsEqualTo(expected);

        await Assert.That(await RenderDocx($"{{{{ Nested.{path} }}}}")).IsEqualTo(expected);
        await Assert.That(await RenderMarkdown($"{{{{ Nested.{path} }}}}")).IsEqualTo(expected);

        await Assert.That(await RenderDocx($"{{% for section in Sections %}}\n\n{{{{ section.{path} }}}}\n\n{{% endfor %}}")).IsEqualTo(expected);
        await Assert.That(await RenderMarkdown($"{{% for section in Sections %}}{{{{ section.{path} }}}}{{% endfor %}}")).IsEqualTo(expected);
    }

    [Test]
    public async Task LoopVariableThatIsATokenStillRenders()
    {
        await Assert.That(await RenderDocx("{% for token in Tokens %}\n\n{{ token }}\n\n{% endfor %}")).IsEqualTo("listed");
        await Assert.That(await RenderMarkdown("{% for token in Tokens %}{{ token }}{% endfor %}")).IsEqualTo("listed");
    }

    static async Task AssertRejected(string docxContent, string markdownContent, string member)
    {
        using var template = DocxTemplateBuilder.Build(docxContent);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains($"Member '{member}' of '");
        await Assert.That(docx.Message).Contains("A token value is rendered whole");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>(markdownContent, styleSource));
        await Assert.That(markdown.Message).Contains($"Member '{member}' of '");
        await Assert.That(markdown.Message).Contains("A token value is rendered whole");
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
