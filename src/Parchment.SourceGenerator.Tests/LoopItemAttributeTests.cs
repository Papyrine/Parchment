// The attribute maps are keyed on dotted paths from the model root through instance members, so a
// token that reaches an annotated member through a loop variable or a static member is an ordinary
// value the runtime hands to Fluid. The rules that belong to a mapped member — alone in its
// paragraph, plain member access — must not be applied to it. The runtime half of the loop cases
// is in Parchment.Tests' LoopItemAttributeTests, over the same templates.
public class LoopItemAttributeTests
{
    const string model =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class HtmlAttribute : System.Attribute { }

        [System.AttributeUsage(System.AttributeTargets.Property)]
        public sealed class MarkdownAttribute : System.Attribute { }

        public class Row
        {
            public string Item { get; set; } = "";
        }

        public class Order
        {
            [ExcelsiorTable]
            public List<Row> Rows { get; set; } = new();

            public List<string> Tags { get; set; } = new();

            [Html]
            public string Notes { get; set; } = "";

            [Markdown]
            public string Memo { get; set; } = "";

            public List<Order> Children { get; set; } = new();
        }

        [ParchmentModel]
        public partial class Invoice
        {
            public List<Order> Orders { get; set; } = new();

            [ExcelsiorTable]
            public List<Row> Rows { get; set; } = new();

            [Html]
            public string Notes { get; set; } = "";

            [ExcelsiorTable]
            public static List<Row> SharedRows { get; set; } = new();

            [Html]
            public static string SharedNotes { get; set; } = "";
        }
        """;

    static readonly string[] attributeRuleIds = ["PARCH007", "PARCH008", "PARCH010", "PARCH024"];

    static readonly string[] loopItemTemplate =
    [
        "{% for order in Orders %}",
        "n={{ order.Rows | size }} t={{ order.Tags | join: '+' }}",
        "{{ order.Notes | upcase }}",
        "{{ order.Notes | html }}",
        "{{ order.Memo | markdown }}",
        "{% for child in order.Children %}",
        "c={{ child.Rows | size }} {{ child.Notes | upcase }}",
        "{% endfor %}",
        "{% endfor %}"
    ];

    static readonly string[] staticTemplate =
    [
        "s={{ SharedRows | size }}",
        "{{ SharedNotes | upcase }}"
    ];

    static readonly string[] rootInsideLoopTemplate =
    [
        "{% for order in Orders %}",
        "n={{ Rows | size }}",
        "{{ Notes | upcase }}",
        "{% endfor %}"
    ];

    static readonly string[] bareFormatTemplate =
    [
        "{% for order in Orders %}",
        "{% for child in order.Children %}",
        "{{ child.Memo }}",
        "{% endfor %}",
        "{% endfor %}"
    ];

    static List<string> Docx(string[] paragraphs) =>
        Ids(GeneratorDriver.Run(model, paragraphs));

    static List<string> Markdown(string[] paragraphs) =>
        Ids(GeneratorDriver.RunMarkdown(model, string.Join("\n\n", paragraphs)));

    static List<string> Ids(GeneratorDriverRunResult result) =>
        result.Results.Single().Diagnostics
            .Select(_ => _.Id)
            .Intersect(attributeRuleIds)
            .Order()
            .ToList();

    [Test]
    public async Task LoopItemMembers_NoAttributeDiagnostics() =>
        await Assert.That(Docx(loopItemTemplate)).IsEmpty();

    [Test]
    public async Task LoopItemMembers_NoAttributeDiagnostics_Markdown() =>
        await Assert.That(Markdown(loopItemTemplate)).IsEmpty();

    [Test]
    public async Task StaticMembers_NoAttributeDiagnostics() =>
        await Assert.That(Docx(staticTemplate)).IsEmpty();

    [Test]
    public async Task StaticMembers_NoAttributeDiagnostics_Markdown() =>
        await Assert.That(Markdown(staticTemplate)).IsEmpty();

    // Inside a loop body a path from the root is still the mapped member, and keeps its rules.
    [Test]
    public async Task RootPathInsideLoop_StillReported() =>
        await Assert.That(Docx(rootInsideLoopTemplate)).IsEquivalentTo(["PARCH007", "PARCH008", "PARCH010"]);

    // PARCH010 is docx-only, and the markdown pass reports one Excelsior rule per token.
    [Test]
    public async Task RootPathInsideLoop_StillReported_Markdown() =>
        await Assert.That(Markdown(rootInsideLoopTemplate)).IsEquivalentTo(["PARCH007"]);

    // The bare path is the one token whose marker is lost without a word.
    [Test]
    public async Task BareFormatMemberThroughNestedLoopVariable_ReportsParch024() =>
        await Assert.That(Docx(bareFormatTemplate)).IsEquivalentTo(["PARCH024"]);

    [Test]
    public async Task BareFormatMemberThroughNestedLoopVariable_ReportsParch024_Markdown() =>
        await Assert.That(Markdown(bareFormatTemplate)).IsEquivalentTo(["PARCH024"]);

    // A condition renders nothing, so there is no markup to show.
    [Test]
    public async Task FormatMemberInCondition_NoParch024_Markdown() =>
        await Assert.That(Markdown(["{% for order in Orders %}", "{% if order.Notes != '' %}x{% endif %}", "{% endfor %}"])).IsEmpty();
}
