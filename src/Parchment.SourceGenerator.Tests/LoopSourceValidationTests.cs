// Build verdicts for what a {% for %} tag names, in both flows. The same templates run through
// runtime registration in Parchment.Tests/LoopSourceValidationTests.cs, and the two files have to
// agree case for case.
public class LoopSourceValidationTests
{
    const string model =
        """
        using System;
        using System.Collections;
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Model
        {
            public string Name { get; set; } = "";
            public IReadOnlyList<string> Lines { get; set; } = [];
            public char[] Letters { get; set; } = [];
            public IEnumerable<char> Chars { get; set; } = [];
            public IEnumerable Bag { get; set; } = new ArrayList();
            public Memory<int> Buffer { get; set; }
            public int Max { get; set; }
            public int Skip { get; set; }
        }
        """;

    static List<string> Docx(string loop) =>
        Ids(GeneratorDriver.Run(model, $"{{% for item in {loop} %}}", "{{ item }}", "{% endfor %}"));

    static List<string> Markdown(string loop) =>
        Ids(GeneratorDriver.RunMarkdown(model, $"{{% for item in {loop} %}}\n{{{{ item }}}}\n\n{{% endfor %}}"));

    static List<string> Ids(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .Where(_ => _.StartsWith("PARCH", StringComparison.Ordinal))
            .ToList();

    [Test]
    [Arguments("Name")]
    [Arguments("Bag")]
    [Arguments("Buffer")]
    public async Task SourceThatIsNotACollection_ReportsParch002(string source)
    {
        await Assert.That(Docx(source)).Contains("PARCH002");
        await Assert.That(Markdown(source)).IsEquivalentTo(["PARCH002"]);
    }

    [Test]
    [Arguments("Letters")]
    [Arguments("Chars")]
    public async Task CharSequence_NoDiagnostic(string source)
    {
        await Assert.That(Docx(source)).IsEmpty();
        await Assert.That(Markdown(source)).IsEmpty();
    }

    [Test]
    [Arguments("Lines limit: Missing")]
    [Arguments("Lines offset: Missing")]
    // The CLR member is not what liquid binds; `size` is.
    [Arguments("Lines limit: Lines.Count")]
    // Evaluated before the loop variable is bound.
    [Arguments("Lines limit: item")]
    [Arguments("(1..Missing)")]
    [Arguments("(1..5) limit: Missing")]
    [Arguments("(1..5) offset: Missing")]
    public async Task UnknownMemberInModifierOrBound_ReportsParch001(string loop)
    {
        await Assert.That(Docx(loop)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(Markdown(loop)).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    [Arguments("Lines limit: Max offset: Skip")]
    [Arguments("Lines limit:2 offset:1 reversed")]
    [Arguments("Lines reversed")]
    [Arguments("Lines limit: Lines.size")]
    [Arguments("(1..Max) offset: Skip")]
    public async Task KnownMemberInModifierOrBound_NoDiagnostic(string loop)
    {
        await Assert.That(Docx(loop)).IsEmpty();
        await Assert.That(Markdown(loop)).IsEmpty();
    }

    // `continue` is liquid's keyword for resuming after the previous loop over the same source,
    // not a model member. Only the markdown flow has the loop state it needs.
    [Test]
    public async Task OffsetContinue_IsAcceptedInMarkdown() =>
        await Assert.That(Markdown("Lines offset: continue")).IsEmpty();
}
