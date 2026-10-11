/// <summary>
/// Generator counterpart of the runtime <c>LiquidMemberTests</c>: the same paths, with the same
/// accept or reject outcome, in both the docx and the markdown flow.
/// </summary>
public class LiquidMemberGeneratorTests
{
    const string model =
        """
        using System.Collections;
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public class Line
        {
            public string Name { get; set; } = "";
            public IReadOnlyList<Line> Children { get; set; } = new List<Line>();
        }

        public class LineBag : IEnumerable<Line>
        {
            public string Label { get; set; } = "";
            public int Count => 0;
            public IEnumerator<Line> GetEnumerator() => null!;
            IEnumerator IEnumerable.GetEnumerator() => null!;
        }

        [ParchmentModel]
        public partial class Doc
        {
            public string Name { get; set; } = "";
            public List<Line> Lines { get; set; } = new();
            public string[] Tags { get; set; } = new string[0];
            public Dictionary<string, string> Lookup { get; set; } = new();
            public LineBag Bag { get; set; } = new();
        }
        """;

    [Test]
    [Arguments("Lines.size")]
    [Arguments("Lines.first.Name")]
    [Arguments("Lines.last.Name")]
    [Arguments("Tags.size")]
    [Arguments("Tags.first")]
    [Arguments("Tags.last")]
    [Arguments("Tags.first.size")]
    [Arguments("Name.size")]
    [Arguments("Name.first")]
    [Arguments("Name.last")]
    [Arguments("Lines.first.Name.size")]
    [Arguments("Lookup.size")]
    [Arguments("Bag.size")]
    [Arguments("Bag.first.Name")]
    public async Task LiquidMemberIsAccepted(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
    }

    [Test]
    [Arguments("Lines.Count")]
    [Arguments("Lines.Capacity")]
    [Arguments("Lines.Size")]
    [Arguments("Lines.First")]
    [Arguments("Lines.first.Nope")]
    [Arguments("Tags.Length")]
    [Arguments("Tags.Rank")]
    [Arguments("Name.Length")]
    [Arguments("Name.Size")]
    [Arguments("Lookup.Count")]
    [Arguments("Lookup.Keys")]
    [Arguments("Lookup.first")]
    [Arguments("Bag.Count")]
    [Arguments("Bag.Label")]
    public async Task ClrMemberIsRejected(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task SizeDrivesAConditional()
    {
        await Assert.That(DocxDiagnostics("{% if Lines.size > 2 %}", "many", "{% endif %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% if Lines.size > 2 %}many{% endif %}")).IsEmpty();
    }

    [Test]
    public async Task FirstIsALoopSource()
    {
        await Assert.That(DocxDiagnostics("{% for child in Lines.first.Children %}", "{{ child.Name }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for child in Lines.first.Children %}{{ child.Name }}{% endfor %}")).IsEmpty();
    }

    [Test]
    public async Task SizeOfALoopVariableMember()
    {
        await Assert.That(DocxDiagnostics("{% for line in Lines %}", "{{ line.Name.size }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for line in Lines %}{{ line.Name.size }}{% endfor %}")).IsEmpty();
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
