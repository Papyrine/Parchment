/// <summary>
/// Generator counterpart of the runtime <c>TokenValueMemberTests</c>: the same paths into a token
/// value, with the same verdict, in both the docx and the markdown flow.
/// </summary>
public class TokenValueMemberGeneratorTests
{
    const string model =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class LabelledToken() : HtmlToken("")
        {
            public string Label { get; set; } = "";
        }

        public class Section
        {
            public HtmlToken Html { get; set; } = new("");
            public MarkdownToken Markdown { get; set; } = new("");
            public OpenXmlToken OpenXml { get; set; } = new();
            public MutateToken Mutate { get; set; } = new();
            public LabelledToken Labelled { get; set; } = new();
            public TokenValue Any { get; set; } = "";
        }

        [ParchmentModel]
        public partial class Doc
        {
            public HtmlToken Html { get; set; } = new("");
            public MarkdownToken Markdown { get; set; } = new("");
            public OpenXmlToken OpenXml { get; set; } = new();
            public MutateToken Mutate { get; set; } = new();
            public LabelledToken Labelled { get; set; } = new();
            public TokenValue Any { get; set; } = "";
            public Section Nested { get; set; } = new();
            public List<Section> Sections { get; set; } = new();
            public List<HtmlToken> Tokens { get; set; } = new();
        }
        """;

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
    public async Task PathIntoATokenIsRejected(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);

        await Assert.That(DocxDiagnostics($"{{{{ Nested.{path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ Nested.{path} }}}}")).IsEquivalentTo(["PARCH001"]);

        await Assert.That(DocxDiagnostics("{% for section in Sections %}", $"{{{{ section.{path} }}}}", "{% endfor %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{% for section in Sections %}}{{{{ section.{path} }}}}{{% endfor %}}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task PathIntoALoopVariableThatIsATokenIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for token in Tokens %}", "{{ token.Source }}", "{% endfor %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics("{% for token in Tokens %}{{ token.Source }}{% endfor %}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task PathIntoATokenInAConditionIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% if Html.Source %}", "some", "{% endif %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics("{% if Html.Source %}some{% endif %}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    [Arguments("Html")]
    [Arguments("Markdown")]
    [Arguments("OpenXml")]
    [Arguments("Mutate")]
    [Arguments("Labelled")]
    [Arguments("Any")]
    public async Task TokenIsAccepted(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEmpty();

        await Assert.That(DocxDiagnostics($"{{{{ Nested.{path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{{{ Nested.{path} }}}}")).IsEmpty();

        await Assert.That(DocxDiagnostics("{% for section in Sections %}", $"{{{{ section.{path} }}}}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{% for section in Sections %}}{{{{ section.{path} }}}}{{% endfor %}}")).IsEmpty();
    }

    [Test]
    public async Task LoopVariableThatIsATokenIsAccepted()
    {
        await Assert.That(DocxDiagnostics("{% for token in Tokens %}", "{{ token }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for token in Tokens %}{{ token }}{% endfor %}")).IsEmpty();
    }

    // No accessor is emitted for a member of a token type: nothing reads one through liquid.
    [Test]
    public async Task NoAccessorIsEmittedForATokenMember()
    {
        var source = string.Join(
            '\n',
            GeneratorDriver.Run(model, ["{{ Html }}"])
                .Results.Single()
                .GeneratedSources
                .Select(_ => _.SourceText.ToString()));
        await Assert.That(source).Contains("\"Html\"");
        await Assert.That(source).DoesNotContain("\"Source\"");
        await Assert.That(source).DoesNotContain("\"Label\"");
        await Assert.That(source).DoesNotContain("\"Render\"");
        await Assert.That(source).DoesNotContain("\"Apply\"");
    }

    static List<string> DocxDiagnostics(params string[] paragraphs) =>
        Ids(GeneratorDriver.Run(model, paragraphs));

    static List<string> MarkdownDiagnostics(string markdown) =>
        Ids(GeneratorDriver.RunMarkdown(model, markdown));

    // PARCH012 is the missing removePersonalInformation warning every test docx trips.
    static List<string> Ids(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .Where(_ => _ != "PARCH012")
            .ToList();
}
