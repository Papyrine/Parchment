/// <summary>
/// Generator counterpart of the runtime <c>SystemTypeMemberTests</c>: the same paths, with the
/// same accept or reject outcome, in both the docx and the markdown flow.
/// </summary>
public class SystemTypeMemberGeneratorTests
{
    const string model =
        """
        using System;
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public struct Meta
        {
            public string Name { get; set; }
        }

        public class Line
        {
            public string Name { get; set; } = "";
        }

        public interface IShape
        {
            string Kind { get; }
        }

        public class LineBase
        {
            public string BaseName { get; set; } = "";
        }

        public class Special : LineBase
        {
        }

        [ParchmentModel]
        public partial class Invoice
        {
            public string Name { get; set; } = "";
            public DateTime When { get; set; }
            public DateTimeOffset Stamp { get; set; }
            public DateOnly Date { get; set; }
            public TimeOnly Time { get; set; }
            public TimeSpan Span { get; set; }
            public Guid Id { get; set; }
            public Uri Link { get; set; } = null!;
            public int Count { get; set; }
            public decimal Amount { get; set; }
            public int? Maybe { get; set; }
            public Meta? Meta { get; set; }
            public (int, string) Pair { get; set; }
            public Tuple<int, string> RefPair { get; set; } = null!;
            public KeyValuePair<string, int> Kv { get; set; }
            public DayOfWeek Day { get; set; }
            public List<DateTime> Dates { get; set; } = new();
            public Dictionary<string, Line> Lookup { get; set; } = new();
            public List<Line> Lines { get; set; } = new();
            public IShape Shape { get; set; } = null!;
            public Special Special { get; set; } = new();
        }
        """;

    [Test]
    [Arguments("When.Year")]
    [Arguments("When.Date")]
    [Arguments("Stamp.Offset")]
    [Arguments("Date.Year")]
    [Arguments("Time.Hour")]
    [Arguments("Span.TotalDays")]
    [Arguments("Id.Version")]
    [Arguments("Link.Host")]
    [Arguments("Count.MaxValue")]
    [Arguments("Amount.Scale")]
    [Arguments("Maybe.Value")]
    [Arguments("Maybe.HasValue")]
    [Arguments("Meta.Value.Name")]
    [Arguments("Meta.Name")]
    [Arguments("Pair.Item1")]
    [Arguments("RefPair.Item1")]
    [Arguments("Day.Friday")]
    [Arguments("Dates.first.Year")]
    [Arguments("Lines.size.MaxValue")]
    public async Task SystemTypeMemberIsRejected(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task SystemTypeMemberOfALoopVariableIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for date in Dates %}", "{{ date.Year }}", "{% endfor %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics("{% for date in Dates %}{{ date.Year }}{% endfor %}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task SystemTypeMemberInAConditionIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% if Maybe.HasValue %}", "some", "{% endif %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics("{% if Maybe.HasValue %}some{% endif %}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task LoopSourceThroughASystemTypeIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for item in RefPair.Item2 %}", "{{ item }}", "{% endfor %}")).Contains("PARCH001");
        await Assert.That(MarkdownDiagnostics("{% for item in RefPair.Item2 %}{{ item }}{% endfor %}")).Contains("PARCH001");
    }

    [Test]
    [Arguments("Maybe")]
    [Arguments("Count")]
    [Arguments("Day")]
    [Arguments("Name.size")]
    [Arguments("Dates.size")]
    [Arguments("Kv.Key")]
    [Arguments("Kv.Value")]
    [Arguments("Lines.first.Name")]
    [Arguments("Shape.Kind")]
    [Arguments("Special.BaseName")]
    public async Task BoundMemberIsAccepted(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
    }

    [Test]
    public async Task MemberOfAUserTypeThroughALoopIsAccepted()
    {
        await Assert.That(DocxDiagnostics("{% for line in Lines %}", "{{ line.Name }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for line in Lines %}{{ line.Name }}{% endfor %}")).IsEmpty();
    }

    [Test]
    public async Task KeyAndValueOfADictionaryLoopAreAccepted()
    {
        await Assert.That(DocxDiagnostics("{% for pair in Lookup %}", "{{ pair.Key }} {{ pair.Value.Name }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for pair in Lookup %}{{ pair.Key }} {{ pair.Value.Name }}{% endfor %}")).IsEmpty();
    }

    [Test]
    [Arguments("model.Name")]
    [Arguments("Model.Name")]
    [Arguments("Invoice.Name")]
    [Arguments("invoice.Name")]
    [Arguments("model")]
    [Arguments("Invoice")]
    public async Task ModelPrefixIsRejected(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task ModelPrefixOnALoopSourceIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for line in model.Lines %}", "{{ line.Name }}", "{% endfor %}")).Contains("PARCH001");
        await Assert.That(MarkdownDiagnostics("{% for line in model.Lines %}{{ line.Name }}{% endfor %}")).Contains("PARCH001");
    }

    [Test]
    public async Task LoopVariableNamedModelIsAccepted()
    {
        await Assert.That(DocxDiagnostics("{% for model in Lines %}", "{{ model.Name }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for model in Lines %}{{ model.Name }}{% endfor %}")).IsEmpty();
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
