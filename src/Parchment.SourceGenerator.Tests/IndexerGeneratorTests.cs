/// <summary>
/// Generator counterpart of the runtime <c>IndexerTests</c>: the same indexed paths, with the same
/// accept or reject outcome, in both the docx and the markdown flow.
/// </summary>
public class IndexerGeneratorTests
{
    const string model =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public sealed class HtmlAttribute : System.Attribute { }

        public sealed class MarkdownAttribute : System.Attribute { }

        public class Line
        {
            public string Name { get; set; } = "";
            public List<string> Tags { get; set; } = new();
        }

        public class Row
        {
            public string Item { get; set; } = "";
            public int Qty { get; set; }
        }

        public class Order
        {
            public List<string> Tags { get; set; } = new();

            [ExcelsiorTable]
            public List<Row> Rows { get; set; } = new();
        }

        public class Budget
        {
            [EditableField]
            public string? Year { get; set; }

            public List<string> Notes { get; set; } = new();
        }

        [ParchmentModel]
        public partial class Doc
        {
            public int Index { get; set; }
            public string Key { get; set; } = "";
            public string[] Tags { get; set; } = new string[0];
            public List<Line> Lines { get; set; } = new();
            public List<List<Line>> Groups { get; set; } = new();
            public Dictionary<string, Line> Lookup { get; set; } = new();

            [ExcelsiorTable]
            public List<Row> Rows { get; set; } = new();

            [Html]
            public string Body { get; set; } = "";

            [Markdown]
            public string Notes { get; set; } = "";

            [EditableField]
            public string Reference { get; set; } = "";

            [EditableField]
            public List<Budget> Budgets { get; set; } = new();

            public Order Order { get; set; } = new();
            public List<Order> Orders { get; set; } = new();
        }
        """;

    // With other text beside it, so that a token taken for the member it indexes would trip the
    // rules that member's own token is held to: PARCH007 for a table, PARCH010 for a filter on a
    // format member, PARCH016 for a filter on an editable one.
    [Test]
    [Arguments("Tags[0]")]
    [Arguments("Tags[1]")]
    [Arguments("Tags[Index]")]
    [Arguments("Tags[0].size")]
    [Arguments("Tags[5]")]
    [Arguments("Order.Tags[0]")]
    [Arguments("Rows[0].Item")]
    [Arguments("Rows[Index].Qty")]
    [Arguments("Rows[0].Item | upcase")]
    [Arguments("Order.Rows[0].Item")]
    [Arguments("Lines[0].Name")]
    [Arguments("Lines[0].Tags[1]")]
    [Arguments("Groups[0][1].Name")]
    [Arguments("Lookup[Key].Name")]
    [Arguments("Body[0]")]
    [Arguments("Body[0] | upcase")]
    [Arguments("Notes[0]")]
    [Arguments("Reference[0] | upcase")]
    public async Task IndexedTokenIsAccepted(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(DocxDiagnostics($"[{{{{ {path} }}}}]")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"[{{{{ {path} }}}}]")).IsEmpty();
    }

    [Test]
    [Arguments("Rows[0].Missing")]
    [Arguments("Rows[Index].Missing")]
    [Arguments("Order.Rows[0].Missing")]
    [Arguments("Lines[0].Missing")]
    [Arguments("Lines[0].Tags[0].Missing")]
    [Arguments("Groups[0][0].Missing")]
    [Arguments("Tags[0].Missing")]
    public async Task UnknownMemberPastAnIndexIsRejected(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task MissingMemberMessageShowsTheIndex()
    {
        var docx = GeneratorDriver.Run(model, "{{ Rows[0].Missing }}");
        await Assert.That(Messages(docx)).IsEquivalentTo(
            ["Template 'Doc.parchment.docx' token '{{ Rows[0].Missing }}' references 'Rows[].Missing' which is not a member of 'Doc'"]);
    }

    [Test]
    public async Task VariableIndexInsideALoop()
    {
        await Assert.That(DocxDiagnostics("{% for i in (0..1) %}", "{{ Tags[i] }}", "{{ Rows[i].Item }};", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for i in (0..1) %}{{ Tags[i] }}{{ Rows[i].Item }};{% endfor %}")).IsEmpty();
    }

    [Test]
    public async Task IndexOnALoopVariableMember()
    {
        await Assert.That(DocxDiagnostics("{% for order in Orders %}", "{{ order.Tags[0] }}", "{{ order.Rows[0].Item }};", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for order in Orders %}{{ order.Tags[0] }}{{ order.Rows[0].Item }};{% endfor %}")).IsEmpty();
    }

    [Test]
    public async Task IndexedLoopSource()
    {
        await Assert.That(
                DocxDiagnostics(
                    "{% for line in Groups[0] %}",
                    "{{ line.Name }};",
                    "{% endfor %}",
                    "{% for tag in Lines[0].Tags %}",
                    "{{ tag.size }}{{ tag }};",
                    "{% endfor %}"))
            .IsEmpty();
        await Assert.That(
                MarkdownDiagnostics(
                    "{% for line in Groups[0] %}{{ line.Name }};{% endfor %}{% for tag in Lines[0].Tags %}{{ tag.size }}{{ tag }};{% endfor %}"))
            .IsEmpty();
    }

    [Test]
    public async Task LoopSourceBehindADictionaryKeyIsUntyped()
    {
        await Assert.That(DocxDiagnostics("{% for tag in Lookup[Key].Tags %}", "{{ tag.size }}{{ tag }};", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics("{% for tag in Lookup[Key].Tags %}{{ tag.size }}{{ tag }};{% endfor %}")).IsEmpty();
    }

    [Test]
    public async Task UnknownMemberPastAnIndexInsideALoopIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for order in Orders %}", "{{ order.Rows[0].Missing }}", "{% endfor %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics("{% for order in Orders %}{{ order.Rows[0].Missing }}{% endfor %}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task UnknownMemberOnAnIndexedLoopSourceIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for line in Groups[0] %}", "{{ line.Missing }}", "{% endfor %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics("{% for line in Groups[0] %}{{ line.Missing }}{% endfor %}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task IndexedLoopSourceThatIsNotEnumerableIsRejected()
    {
        await Assert.That(DocxDiagnostics("{% for part in Lines[0] %}", "x", "{% endfor %}")).IsEquivalentTo(["PARCH002"]);
        await Assert.That(MarkdownDiagnostics("{% for part in Lines[0] %}x{% endfor %}")).IsEquivalentTo(["PARCH002"]);
    }

    // Twice in the body would be PARCH017 for the field itself, and a loop over the collection
    // with a modifier PARCH027.
    [Test]
    public async Task IndexedEditableMemberIsNotTheField() =>
        await Assert.That(
                DocxDiagnostics(
                    "[{{ Reference[0] }}]",
                    "[{{ Reference[0] }}]",
                    "{% for note in Budgets[0].Notes limit:1 %}",
                    "{{ note }};",
                    "{% endfor %}"))
            .IsEmpty();

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

    static List<string> Messages(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Where(_ => _.Id != "PARCH012")
            .Select(_ => _.GetMessage())
            .ToList();
}
