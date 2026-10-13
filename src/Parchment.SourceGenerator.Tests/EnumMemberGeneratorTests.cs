/// <summary>
/// Generator counterpart of the runtime <c>EnumMemberTests</c>: the same paths, with the same
/// accept or reject outcome, in both the docx and the markdown flow.
/// </summary>
public class EnumMemberGeneratorTests
{
    const string model =
        """
        using System;
        using System.Collections.Generic;
        using Parchment;

        namespace Sample;

        public enum Status
        {
            Draft,
            Final
        }

        [Flags]
        public enum Access
        {
            Read = 1,
            Write = 2
        }

        public class Holder
        {
            public Status Status { get; set; }
        }

        [ParchmentModel]
        public partial class Doc
        {
            public Status Status { get; set; }
            public Status? Maybe { get; set; }
            public Access Access { get; set; }
            public Phase Stage { get; set; }
            public DayOfWeek Day { get; set; }
            public List<Status> History { get; set; } = new();
            public Holder Holder { get; set; } = new();

            public enum Phase
            {
                Open,
                Done
            }
        }
        """;

    [Test]
    [Arguments("Status")]
    [Arguments("Maybe")]
    [Arguments("Access")]
    [Arguments("Stage")]
    [Arguments("Day")]
    [Arguments("Holder.Status")]
    public async Task EnumIsAccepted(string path)
    {
        await Assert.That(Ids(Docx($"{{{{ {path} }}}}"))).IsEmpty();
        await Assert.That(Ids(Markdown($"{{{{ {path} }}}}"))).IsEmpty();
    }

    [Test]
    public async Task ComparingAgainstTheTextIsAccepted()
    {
        await Assert.That(Ids(Docx("""{% if Status == "Final" %}""", "final", "{% else %}", "not final", "{% endif %}"))).IsEmpty();
        await Assert.That(Ids(Markdown("""{% if Status == "Final" %}final{% else %}not final{% endif %}"""))).IsEmpty();
    }

    [Test]
    [Arguments("Status.Final", "Status")]
    [Arguments("Status.Draft", "Status")]
    [Arguments("Status.value__", "Status")]
    [Arguments("Maybe.Final", "Maybe")]
    [Arguments("Maybe.Value", "Maybe")]
    [Arguments("Maybe.Value.Final", "Maybe")]
    [Arguments("Access.Write", "Access")]
    [Arguments("Stage.Done", "Stage")]
    [Arguments("Day.Monday", "Day")]
    [Arguments("Holder.Status.Final", "Holder.Status")]
    public async Task ConstantIsRejectedNamingTheEnum(string path, string enumPath)
    {
        var docx = Docx($"{{{{ {path} }}}}");
        await Assert.That(Ids(docx)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(docx.Single().GetMessage()).Contains($"references '{path}', but '{enumPath}' is an enum");

        var markdown = Markdown($"{{{{ {path} }}}}");
        await Assert.That(Ids(markdown)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(markdown.Single().GetMessage()).Contains($"references '{path}', but '{enumPath}' is an enum");
    }

    // The reported template: clean under the generator, then "not final" for a Final status.
    [Test]
    public async Task ConstantInAConditionIsRejected()
    {
        var docx = Docx("{% if Status == Status.Final %}", "final", "{% else %}", "not final", "{% endif %}");
        await Assert.That(Ids(docx)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(docx.Single().GetMessage()).Contains("'Status' is an enum");

        var markdown = Markdown("{% if Status == Status.Final %}final{% else %}not final{% endif %}");
        await Assert.That(Ids(markdown)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(markdown.Single().GetMessage()).Contains("'Status' is an enum");
    }

    [Test]
    public async Task ConstantThroughALoopVariableIsRejected()
    {
        var docx = Docx("{% for entry in History %}", "{{ entry.Final }}", "{% endfor %}");
        await Assert.That(Ids(docx)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(docx.Single().GetMessage()).Contains("'entry' is an enum");

        var markdown = Markdown("{% for entry in History %}{{ entry.Final }}{% endfor %}");
        await Assert.That(Ids(markdown)).IsEquivalentTo(["PARCH001"]);
        await Assert.That(markdown.Single().GetMessage()).Contains("'entry' is an enum");
    }

    // A member that does not exist keeps the plain wording.
    [Test]
    public async Task UnknownMemberBesideAnEnumKeepsThePlainMessage()
    {
        var docx = Docx("{{ Holder.Nope }}");
        await Assert.That(docx.Single().GetMessage()).Contains("references 'Holder.Nope' which is not a member of 'Doc'");

        var markdown = Markdown("{{ Holder.Nope }}");
        await Assert.That(markdown.Single().GetMessage()).Contains("references 'Holder.Nope' which is not a member of 'Doc'");
    }

    // Roslyn lists an enum's constants as public static fields, which is all an accessor needs.
    // Nothing reads one: the value never reaches Fluid as an object.
    [Test]
    public async Task NoAccessorsAreEmittedForAnEnum()
    {
        var generated = GeneratorDriver.Run(model, "{{ Status }}")
            .Results.Single()
            .GeneratedSources.Single()
            .SourceText.ToString();

        await Assert.That(generated).Contains("_Accessors_global__Sample_Doc");
        await Assert.That(generated).Contains("_Accessors_global__Sample_Holder");
        await Assert.That(generated).DoesNotContain("_Accessors_global__Sample_Status");
        await Assert.That(generated).DoesNotContain("_Accessors_global__Sample_Access");
        await Assert.That(generated).DoesNotContain("_Accessors_global__Sample_Doc_Phase");
        await Assert.That(generated).DoesNotContain("Status.Draft");
    }

    static List<Diagnostic> Docx(params string[] paragraphs) =>
        Report(GeneratorDriver.Run(model, paragraphs));

    static List<Diagnostic> Markdown(string markdown) =>
        Report(GeneratorDriver.RunMarkdown(model, markdown));

    // PARCH012 is the missing removePersonalInformation warning every test docx trips.
    static List<Diagnostic> Report(GeneratorDriverRunResult result) =>
        result.Results.Single()
            .Diagnostics
            .Where(_ => _.Id != "PARCH012")
            .ToList();

    static List<string> Ids(List<Diagnostic> diagnostics) =>
        diagnostics
            .Select(_ => _.Id)
            .ToList();
}
