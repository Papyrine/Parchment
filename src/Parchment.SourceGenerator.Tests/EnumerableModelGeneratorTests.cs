/// <summary>
/// Generator counterpart of the runtime <c>EnumerableModelTests</c>: a model that is itself a
/// collection binds what it declares at the root, and is an array everywhere else.
/// </summary>
public class EnumerableModelGeneratorTests
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
        }

        [ParchmentModel]
        public partial class Order : List<Line>
        {
            public string Name { get; set; } = "";
            public IEnumerable<Line> Items => this;
            public Order? Previous { get; set; }
            public List<Order> History { get; set; } = new();
        }
        """;

    const string dictionaryModel =
        """
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        [ParchmentModel]
        public partial class Lookup : Dictionary<string, string>
        {
            public string Name { get; set; } = "";
        }
        """;

    [Test]
    [Arguments("Name")]
    [Arguments("Items.size")]
    [Arguments("Items.first.Name")]
    [Arguments("Previous.size")]
    [Arguments("Previous.last.Name")]
    [Arguments("History.first.size")]
    [Arguments("History.first.last.Name")]
    public async Task IsAccepted(string path)
    {
        await Assert.That(DocxDiagnostics(model, $"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics(model, $"{{{{ {path} }}}}")).IsEmpty();
    }

    [Test]
    [Arguments("Count")]
    [Arguments("Capacity")]
    [Arguments("size")]
    [Arguments("first")]
    [Arguments("last")]
    [Arguments("Previous.Name")]
    [Arguments("Previous.Count")]
    [Arguments("History.first.Name")]
    public async Task IsRejected(string path)
    {
        await Assert.That(DocxDiagnostics(model, $"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics(model, $"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
    }

    [Test]
    public async Task DictionaryModelBindsItsOwnMember()
    {
        await Assert.That(DocxDiagnostics(dictionaryModel, "{{ Name }}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics(dictionaryModel, "{{ Name }}")).IsEmpty();
    }

    [Test]
    public async Task CollectionMemberOfADictionaryModelIsNotALoopSource()
    {
        await Assert.That(DocxDiagnostics(dictionaryModel, "{% for key in Keys %}", "{{ key }}", "{% endfor %}")).IsNotEmpty();
        await Assert.That(MarkdownDiagnostics(dictionaryModel, "{% for key in Keys %}{{ key }}{% endfor %}")).IsNotEmpty();
    }

    [Test]
    public async Task ItemsAreLoopedThroughAMember()
    {
        await Assert.That(DocxDiagnostics(model, "{% for line in Items %}", "{{ line.Name }}", "{% endfor %}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics(model, "{% for line in Items %}{{ line.Name }}{% endfor %}")).IsEmpty();
    }

    [Test]
    public async Task LoopElementOfTheModelTypeIsAnArray()
    {
        await Assert.That(
                DocxDiagnostics(
                    model,
                    "{% for order in History %}",
                    "{{ order.size }}{{ order.first.Name }}{{ order.last.Name }}",
                    "{% for line in order %}",
                    "{{ line.Name }}",
                    "{% endfor %}",
                    "{% endfor %}"))
            .IsEmpty();
        await Assert.That(
                MarkdownDiagnostics(
                    model,
                    "{% for order in History %}{{ order.size }}{{ order.first.Name }}{% for line in order %}{{ line.Name }}{% endfor %}{% endfor %}"))
            .IsEmpty();
    }

    [Test]
    [Arguments("Name")]
    [Arguments("Count")]
    public async Task OwnMemberOfALoopElementIsRejected(string member)
    {
        await Assert.That(DocxDiagnostics(model, "{% for order in History %}", $"{{{{ order.{member} }}}}", "{% endfor %}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics(model, $"{{% for order in History %}}{{{{ order.{member} }}}}{{% endfor %}}")).IsEquivalentTo(["PARCH001"]);
    }

    static List<string> DocxDiagnostics(string source, params string[] paragraphs) =>
        Ids(GeneratorDriver.Run(source, paragraphs));

    static List<string> MarkdownDiagnostics(string source, string markdown) =>
        Ids(GeneratorDriver.RunMarkdown(source, markdown));

    // PARCH012 is the missing removePersonalInformation warning every test docx trips.
    static List<string> Ids(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Select(_ => _.Id)
            .Where(_ => _ != "PARCH012")
            .ToList();
}
