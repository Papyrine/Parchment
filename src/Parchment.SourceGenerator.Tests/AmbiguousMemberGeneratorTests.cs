/// <summary>
/// Generator counterpart of the runtime <c>AmbiguousMemberTests</c>: a name more than one member
/// answers to gets one accessor, for the most derived and then the first declared member, and
/// every spelling of it validates against that member.
/// </summary>
public class AmbiguousMemberGeneratorTests
{
    const string model =
        """
        using Parchment;

        namespace Sample;

        public class Root
        {
            public decimal Total { get; set; }
            public decimal Tax { get; set; }
            public decimal Rate;
            public static decimal Shared => 4m;
        }

        public class Middle : Root
        {
            public new string Tax { get; set; } = "";
        }

        public class Leaf : Middle
        {
            public new string Total { get; set; } = "";
            public new string Rate = "";
            public new static string Shared => "";
        }

        public class Cased
        {
            public string Name { get; set; } = "";
            public int name { get; set; }
            public string Code = "";
            public int CODE;
            public string Total { get; set; } = "";
            public int total;
            public string count = "";
            public int Count { get; set; }
        }

        [ParchmentModel]
        public partial class Doc
        {
            public Leaf Leaf { get; set; } = new();
            public Cased Cased { get; set; } = new();
        }
        """;

    [Test]
    [Arguments("Leaf.Total.size")]
    [Arguments("Leaf.Tax.size")]
    [Arguments("Leaf.Rate.size")]
    [Arguments("Leaf.Shared.size")]
    [Arguments("Cased.Name.size")]
    [Arguments("Cased.name.size")]
    [Arguments("Cased.CODE.size")]
    [Arguments("Cased.total.size")]
    [Arguments("Cased.Count.size")]
    public async Task EverySpellingIsTheFirstDeclaredMember(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEmpty();
    }

    [Test]
    [Arguments("Leaf.Total.Nope")]
    [Arguments("Leaf.Tax.Nope")]
    [Arguments("Cased.name.Nope")]
    public async Task MissingMemberIsReported(string path)
    {
        await Assert.That(DocxDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
        await Assert.That(MarkdownDiagnostics($"{{{{ {path} }}}}")).IsEquivalentTo(["PARCH001"]);
    }

    // One accessor per name, so the generated table never holds two keys a case-insensitive
    // lookup would collide on. That the accessors compile is covered by the runtime tests, whose
    // models are the same shapes built for real.
    [Test]
    public async Task OneAccessorPerName()
    {
        var emitted = string.Concat(
            GeneratorDriver.Run(model, "{{ Cased.Name }}")
                .Results.SelectMany(_ => _.GeneratedSources)
                .Select(_ => _.SourceText.ToString()));

        await Assert.That(emitted).Contains("""new("Name", """);
        await Assert.That(emitted).DoesNotContain("""new("name", """);
        await Assert.That(emitted).Contains("""new("Code", """);
        await Assert.That(emitted).DoesNotContain("""new("CODE", """);
        await Assert.That(emitted).Contains("""new("count", """);
        await Assert.That(emitted).DoesNotContain("""new("Count", """);
        await Assert.That(emitted).Contains("((global::Sample.Leaf)o).Total");
        await Assert.That(emitted).DoesNotContain("total)");
        await Assert.That(Occurrences(emitted, """new("Total", """)).IsEqualTo(2);
        await Assert.That(Occurrences(emitted, """new("Tax", """)).IsEqualTo(1);
    }

    static int Occurrences(string text, string value) =>
        text.Split(value).Length - 1;

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
