// PARCH037 — the compile-time half of the runtime's "a filter that produces a token is the last
// filter of its token" check (FilterTests and EscapingTests in Parchment.Tests). Which filters
// produce a token depends on the flow, so each case runs against the flow it names.
public class FilterOrderGeneratorTests
{
    const string source =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Note
        {
            public string Title { get; set; } = "";

            public string Body { get; set; } = "";

            public List<string> Items { get; set; } = new();
        }
        """;

    [Test]
    [Arguments("{{ Body | markdown | upcase }}", "markdown", "upcase")]
    [Arguments("{{ Body | html | upcase }}", "html", "upcase")]
    [Arguments("{{ Items | bullet_list | upcase }}", "bullet_list", "upcase")]
    [Arguments("{{ Items | numbered_list | append: 'x' }}", "numbered_list", "append")]
    [Arguments("{{ Body | markdown | raw }}", "markdown", "raw")]
    [Arguments("{{ Body | markdown | html }}", "markdown", "html")]
    [Arguments("{{ Body | upcase | markdown | downcase | strip }}", "markdown", "downcase")]
    public async Task Docx_FilterAfterTokenFilter_ReportsParch037(string token, string tokenFilter, string following)
    {
        var diagnostics = GeneratorDriver.Run(source, $"Before {token} after").Results.Single().Diagnostics;

        var diagnostic = diagnostics.Single();
        var message = diagnostic.GetMessage();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH037");
        await Assert.That(diagnostic.Severity).IsEqualTo(DiagnosticSeverity.Error);
        await Assert.That(message).Contains($"token '{token}'");
        await Assert.That(message).Contains($"Filter '{following}' is applied after '{tokenFilter}'");
        await Assert.That(message).Contains($"Move '{following}' before '{tokenFilter}'");
    }

    // Reported wherever the token sits: the paragraph list covers every part.
    [Test]
    public async Task Docx_FilterAfterTokenFilterInsideLoop_ReportsParch037()
    {
        var diagnostics = GeneratorDriver.Run(
                source,
                "{% for item in Items %}",
                "{{ item | markdown | upcase }}",
                "{% endfor %}")
            .Results.Single().Diagnostics;

        await Assert.That(diagnostics.Select(_ => _.Id).ToList()).IsEquivalentTo(["PARCH037"]);
    }

    [Test]
    [Arguments("{{ Body | upcase | markdown }}")]
    [Arguments("{{ Body | append: Title | upcase | html }}")]
    [Arguments("{{ Items | sort | bullet_list }}")]
    [Arguments("{{ Items | sort | numbered_list }}")]
    [Arguments("{{ Body | escape_xml | upcase }}")]
    [Arguments("{{ Body | markdown }}")]
    public async Task Docx_TokenFilterLast_NoDiagnostic(string token)
    {
        var diagnostics = GeneratorDriver.Run(source, token).Results.Single().Diagnostics;

        await Assert.That(diagnostics).IsEmpty();
    }

    // A markdown template holds the parsed template and not the token's text, so the token is
    // named by its chain.
    [Test]
    [Arguments("{{ Body | html | upcase }}", "upcase", "Body | html | upcase")]
    [Arguments("{{ Body | html | append: ' tail' }}", "append", "Body | html | append")]
    [Arguments("{{ Body | upcase | html | downcase | strip }}", "downcase", "Body | upcase | html | downcase | strip")]
    [Arguments("{% assign shout = Body | html | upcase %}{{ shout }}", "upcase", "Body | html | upcase")]
    [Arguments("{% capture held %}{{ Body | html | upcase }}{% endcapture %}{{ held }}", "upcase", "Body | html | upcase")]
    [Arguments("{% for item in Items %}{{ item | html | upcase }}{% endfor %}", "upcase", "item | html | upcase")]
    [Arguments("{% if Title %}x{% else %}{{ Body | html | upcase }}{% endif %}", "upcase", "Body | html | upcase")]
    public async Task Markdown_FilterAfterHtmlFilter_ReportsParch037(string markdown, string following, string chain)
    {
        var diagnostics = GeneratorDriver.RunMarkdown(source, markdown).Results.Single().Diagnostics;

        var diagnostic = diagnostics.Single();
        var message = diagnostic.GetMessage();
        await Assert.That(diagnostic.Id).IsEqualTo("PARCH037");
        await Assert.That(diagnostic.Severity).IsEqualTo(DiagnosticSeverity.Error);
        await Assert.That(message).Contains($"token '{chain}'");
        await Assert.That(message).Contains($"Filter '{following}' is applied after 'html'");
        await Assert.That(message).Contains($"Move '{following}' before 'html'");
    }

    // In a markdown template the other three write markdown source, which a later filter can still
    // work on, and the runtime renders these. The same chains are PARCH037 in a docx template.
    [Test]
    [Arguments("{{ Body | markdown | upcase }}")]
    [Arguments("{{ Items | bullet_list | upcase }}")]
    [Arguments("{{ Items | numbered_list | upcase }}")]
    [Arguments("{{ Body | upcase | html }}")]
    [Arguments("{{ Body | markdown | html }}")]
    [Arguments("{{ Body | html }}")]
    public async Task Markdown_NoTokenFilterBeforeAnother_NoDiagnostic(string markdown)
    {
        var diagnostics = GeneratorDriver.RunMarkdown(source, markdown).Results.Single().Diagnostics;

        await Assert.That(diagnostics).IsEmpty();
    }
}
