using System.IO.Compression;

/// <summary>
/// How a <c>[ParchmentModel]</c> type finds its files by convention: the template is the
/// AdditionalFile named <c>TypeName.parchment.md</c> or <c>TypeName.parchment.docx</c>, and a
/// markdown template's style source is the <c>TypeName.parchment.dotx</c> when the project carries
/// one, otherwise the nearest <c>parchment.dotx</c> up the directory tree.
/// </summary>
public class TemplateConventionTests
{
    const string reportModel =
        """
        using Parchment;

        [ParchmentModel]
        public partial class Report
        {
            public required string Name { get; init; }
        }
        """;

    static GeneratorDriverRunResult Run(params TemplateFile[] files)
    {
        var setup = GeneratorDriver.CreateDriverWithFiles(reportModel, files);
        return setup.Driver.RunGenerators(setup.Compilation).GetRunResult();
    }

    static IEnumerable<string> Codes(GeneratorDriverRunResult result) =>
        result.Diagnostics.Select(_ => _.Id);

    static string Generated(GeneratorDriverRunResult result) =>
        string.Concat(result.GeneratedTrees.Select(_ => _.GetText().ToString()));

    static byte[] Markdown(string text) => Encoding.UTF8.GetBytes(text);

    [Test]
    public async Task TemplateBesideNothingElse()
    {
        var result = Run(new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    // The convention matches by file name wherever the file sits — templates in a Templates/
    // folder need no path in the attribute.
    [Test]
    public async Task TemplateInANestedFolder()
    {
        var result = Run(new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task MatchIsCaseInsensitive()
    {
        var result = Run(new TemplateFile("report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    [Test]
    public async Task NoMatchIsReported()
    {
        var result = Run(new TemplateFile("Other.parchment.md", Markdown("Hello")));

        await Assert.That(Codes(result)).Contains("PARCH004");
    }

    static string Messages(GeneratorDriverRunResult result) =>
        string.Join('\n', result.Diagnostics.Select(_ => _.GetMessage()));

    // The list is what separates "named wrong" from "nothing arrived", and the near-miss is
    // usually sitting in it.
    //
    // Run through RunInProject so ProjectDir is set and the paths come out project-relative, the
    // way a real build reports them — a bare file name would not tell two namesakes apart.
    [Test]
    public async Task NoMatchNamesTheTemplatesItDidSee()
    {
        var result = GeneratorDriver.RunInProject(
            reportModel,
            new TemplateFile("Templates/Other.parchment.md", Markdown("Hello")),
            new TemplateFile("Templates/Reprot.parchment.docx", GeneratorDriver.BuildDocxBytes("typo")));

        await Assert.That(Messages(result)).Contains(
            "Templates seen: Templates/Other.parchment.md, Templates/Reprot.parchment.docx.");
    }

    // The other shape: no template reached the generator, which is a build or IDE problem rather
    // than a naming one. Renaming anything would be wasted effort, so the message says so.
    [Test]
    public async Task NoMatchWithNoTemplatesAtAllSaysSo()
    {
        var result = Run();

        await Assert.That(Messages(result)).Contains("No templates reached the generator at all.");
    }

    // An unmarked namesake is not a template, so it must not turn up in the list either — it would
    // read as "your template is right there" when the fix is to rename it.
    [Test]
    public async Task TemplatesSeenExcludesUnmarkedFiles()
    {
        var result = Run(new TemplateFile("Docs/Report.md", Markdown("Notes")));

        await Assert.That(Messages(result)).Contains("No templates reached the generator at all.");
    }

    // The marker is the whole reason the package can glob templates in without a csproj entry:
    // an unmarked file is not a template, however it is named and however it is declared.
    [Test]
    public async Task UnmarkedFileNamedAfterTheTypeDoesNotBind()
    {
        var result = Run(new TemplateFile("Report.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH004");
    }

    // The case that motivated the marker: a design note sitting beside the template, named after
    // the model because it is about the model. It must not become a second candidate.
    [Test]
    public async Task UnmarkedNamesakeIsNotAmbiguousWithTheTemplate()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Docs/Report.md", Markdown("Notes about the report model.")));

        await Assert.That(Codes(result)).IsEmpty();
    }

    // An unmarked dotx is not a style document either — including one named after the type.
    [Test]
    public async Task UnmarkedDotxIsNotAStyleSource()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var result = Run(
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Report.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // The shared style document keeps its bare name — parchment.dotx, not parchment.parchment.dotx
    // — so a model actually named Parchment is the one case where the two could be confused. The
    // marker keeps them apart.
    [Test]
    public async Task ModelNamedParchmentDoesNotBindTheSharedStyleDoc()
    {
        var sharedBytes = GeneratorDriver.BuildDocxBytes("shared");
        // Namespaced: a type named Parchment in the global namespace would make `using Parchment;`
        // a CS0138, and the model would never reach the generator to prove anything.
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class Parchment
            {
                public required string Name { get; init; }
            }
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(
            source,
            new TemplateFile("Templates/Parchment.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/parchment.dotx", sharedBytes));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();

        // Not PARCH021: the shared document is not a TypeName match, it is the folder's default.
        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(sharedBytes));
    }

    // Two namesakes in different folders: neither is preferred.
    [Test]
    public async Task NamesakesInDifferentFoldersAreAmbiguous()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Drafts/Report.parchment.md", Markdown("Hello {{ Name }}")));

        await Assert.That(Codes(result)).Contains("PARCH020");
    }

    // AdditionalFiles order follows the csproj, the globs and the file system, none of which the
    // model's author changed. Neither the diagnostics nor the generated source may follow it.
    static async Task AssertSameWhateverTheFileOrder(string source, params TemplateFile[] files)
    {
        var forward = GeneratorDriver.RunInProject(source, files);
        var reversed = GeneratorDriver.RunInProject(source, files.AsEnumerable().Reverse().ToArray());

        await Assert.That(Describe(reversed)).IsEqualTo(Describe(forward));
        await Assert.That(Generated(reversed)).IsEqualTo(Generated(forward));
    }

    // Id, where it points and what it says. The path is reduced to the file name: each run gets a
    // project directory of its own.
    static string Describe(GeneratorDriverRunResult result)
    {
        var builder = new StringBuilder();
        foreach (var diagnostic in result.Diagnostics)
        {
            var span = diagnostic.Location.GetLineSpan();
            builder.AppendLine(
                $"{diagnostic.Id} {Path.GetFileName(span.Path)}{span.Span} +{diagnostic.AdditionalLocations.Count}: {diagnostic.GetMessage()}");
        }

        return builder.ToString();
    }

    [Test]
    public async Task AmbiguousTemplatesAreListedInPathOrder()
    {
        var result = GeneratorDriver.RunInProject(
            reportModel,
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Drafts/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Archive/Report.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }}")));

        await Assert.That(Messages(result)).Contains(
            "more than one template file: Archive/Report.parchment.docx, Drafts/Report.parchment.md, Templates/Report.parchment.md.");
    }

    [Test]
    public async Task AmbiguousStyleDocsAreListedInPathOrder()
    {
        var result = GeneratorDriver.RunInProject(
            reportModel,
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")),
            new TemplateFile("Drafts/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Messages(result)).Contains(
            "more than one style document: Drafts/Report.parchment.dotx, Templates/Report.parchment.dotx.");
    }

    [Test]
    public Task AmbiguousTemplateIsTheSameWhateverTheFileOrder() =>
        AssertSameWhateverTheFileOrder(
            reportModel,
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Drafts/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Archive/Report.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }}")));

    [Test]
    public Task AmbiguousStyleDocIsTheSameWhateverTheFileOrder() =>
        AssertSameWhateverTheFileOrder(
            reportModel,
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")),
            new TemplateFile("Drafts/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

    [Test]
    public Task MissingTemplateIsTheSameWhateverTheFileOrder() =>
        AssertSameWhateverTheFileOrder(
            reportModel,
            new TemplateFile("Templates/Other.parchment.md", Markdown("Hello")),
            new TemplateFile("Templates/Reprot.parchment.docx", GeneratorDriver.BuildDocxBytes("typo")),
            new TemplateFile("Another.parchment.md", Markdown("Hello")));

    // Several models, each with diagnostics of its own, and a style document picked from more than
    // one candidate: what is reported, in what order, and what is generated all hold.
    [Test]
    public Task SeveralModelsAreTheSameWhateverTheFileOrder()
    {
        var source =
            """
            using Parchment;

            [ParchmentModel]
            public partial class Report
            {
                public required string Name { get; init; }
            }

            [ParchmentModel]
            public partial class Letter
            {
                public required string Name { get; init; }
            }

            [ParchmentModel]
            public partial class Memo
            {
                public required string Name { get; init; }
            }
            """;

        return AssertSameWhateverTheFileOrder(
            source,
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }} {{ Missing }} {{ Gone }}")),
            new TemplateFile("Templates/parchment.dotx", GeneratorDriver.BuildDocxBytes("near")),
            new TemplateFile("parchment.dotx", GeneratorDriver.BuildDocxBytes("far")),
            new TemplateFile("Letter.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }} {{ Missing }}", "{{ Gone }}")),
            new TemplateFile("Other/Letter.parchment.dotx", GeneratorDriver.BuildDocxBytes("unused")));
    }

    [Test]
    public async Task TypeNamedDotxIsTheStyleSource()
    {
        var result = Run(
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).IsEmpty();
        // The style source is embedded as the second registration argument.
        await Assert.That(Generated(result)).Contains("RegisterMarkdownTemplate");
        await Assert.That(Generated(result)).Contains("EmbeddedBytes(\"");
    }

    [Test]
    public async Task TwoTypeNamedDotxesAreAmbiguous()
    {
        var result = Run(
            new TemplateFile("Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")),
            new TemplateFile("Drafts/Report.parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).Contains("PARCH021");
    }

    // parchment.dotx beside the template applies when no TypeName.dotx exists.
    [Test]
    public async Task SharedParchmentDotxApplies()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/parchment.dotx", GeneratorDriver.BuildDocxBytes("styled")));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains("EmbeddedBytes(\"");
    }

    // The nearest ancestor's parchment.dotx wins over one further up.
    [Test]
    public async Task NearestParchmentDotxWins()
    {
        var nearBytes = GeneratorDriver.BuildDocxBytes("near");
        var farBytes = GeneratorDriver.BuildDocxBytes("far");
        var result = Run(
            new TemplateFile("Templates/Inner/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Inner/parchment.dotx", nearBytes),
            new TemplateFile("parchment.dotx", farBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(nearBytes));
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(farBytes));
    }

    // A parchment.dotx in a sibling folder is not on the template's ancestor path and must not
    // apply.
    [Test]
    public async Task SiblingFolderParchmentDotxDoesNotApply()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Other/parchment.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // TypeName.dotx wins over a shared parchment.dotx, even one beside the template.
    [Test]
    public async Task TypeNamedDotxWinsOverSharedParchmentDotx()
    {
        var typeBytes = GeneratorDriver.BuildDocxBytes("typed");
        var sharedBytes = GeneratorDriver.BuildDocxBytes("shared");
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Templates/Report.parchment.dotx", typeBytes),
            new TemplateFile("Templates/parchment.dotx", sharedBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(typeBytes));
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(sharedBytes));
    }

    // A style source the runtime could not register fails the build, as an unreadable docx
    // template does, and at the file rather than at the model.
    [Test]
    [MethodDataSource(nameof(UnusableStyleDocs))]
    public async Task UnusableStyleDocIsReported(string name, byte[] bytes)
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile(name, bytes));

        await Assert.That(Codes(result)).IsEquivalentTo(["PARCH006"]);
        var diagnostic = result.Diagnostics.Single();
        await Assert.That(Path.GetFileName(diagnostic.Location.GetLineSpan().Path)).IsEqualTo(Path.GetFileName(name));
        await Assert.That(diagnostic.GetMessage()).Contains(Path.GetFileName(name));
        await Assert.That(Generated(result)).IsEmpty();
    }

    public static IEnumerable<Func<(string, byte[])>> UnusableStyleDocs()
    {
        var good = GeneratorDriver.BuildDocxBytes("styled");
        yield return () => ("Templates/parchment.dotx", "this is not a zip"u8.ToArray());
        yield return () => ("Templates/parchment.dotx", []);
        yield return () => ("parchment.dotx", good[..(good.Length / 2)]);
        yield return () => ("Templates/parchment.dotx", Repack(good, static name => name != "word/document.xml", static text => text));
        yield return () => ("Templates/parchment.dotx", Repack(good, static _ => true, static text => text.Replace("</w:body>", "")));
        yield return () => ("Templates/Report.parchment.dotx", "this is not a zip"u8.ToArray());
    }

    static byte[] Repack(byte[] docx, Func<string, bool> keep, Func<string, string> rewrite)
    {
        using var input = new ZipArchive(new MemoryStream(docx), ZipArchiveMode.Read);
        using var stream = new MemoryStream();
        using (var output = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            foreach (var entry in input.Entries.Where(_ => keep(_.FullName)))
            {
                using var reader = new StreamReader(entry.Open());
                using var writer = new StreamWriter(output.CreateEntry(entry.FullName).Open());
                writer.Write(rewrite(reader.ReadToEnd()));
            }
        }

        return stream.ToArray();
    }

    // One file, one error: sixty models sharing a parchment.dotx is the ordinary case.
    [Test]
    public async Task UnusableSharedStyleDocIsReportedOncePerFile()
    {
        var source =
            """
            using Parchment;

            namespace Sample;

            [ParchmentModel]
            public partial class First;

            [ParchmentModel]
            public partial class Second;

            [ParchmentModel]
            public partial class Third;
            """;
        var setup = GeneratorDriver.CreateDriverWithFiles(
            source,
            new TemplateFile("First.parchment.md", Markdown("one")),
            new TemplateFile("Second.parchment.md", Markdown("two")),
            new TemplateFile("Other/Third.parchment.md", Markdown("three")),
            new TemplateFile("parchment.dotx", "this is not a zip"u8.ToArray()),
            new TemplateFile("Other/parchment.dotx", "nor is this"u8.ToArray()));
        var result = setup.Driver.RunGenerators(setup.Compilation).GetRunResult();

        await Assert.That(Codes(result)).IsEquivalentTo(["PARCH006", "PARCH006"]);
        await Assert.That(result.Diagnostics.Select(_ => _.Location.GetLineSpan().Path).Distinct().Count()).IsEqualTo(2);
    }

    // An unusable style document nothing inherits is not this build's problem.
    [Test]
    public async Task UnusableStyleDocNothingUsesIsNotReported()
    {
        var result = Run(
            new TemplateFile("Templates/Report.parchment.md", Markdown("Hello {{ Name }}")),
            new TemplateFile("Other/parchment.dotx", "this is not a zip"u8.ToArray()));

        await Assert.That(Codes(result)).IsEmpty();
    }

    // Docx templates carry their own styles; a dotx present in the project must not affect them.
    [Test]
    public async Task DocxTemplateIgnoresStyleDocs()
    {
        var styleBytes = GeneratorDriver.BuildDocxBytes("styled");
        var result = Run(
            new TemplateFile("Report.parchment.docx", GeneratorDriver.BuildDocxBytes("Hello {{ Name }}")),
            new TemplateFile("Report.parchment.dotx", styleBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains("RegisterDocxTemplate");
        await Assert.That(Generated(result)).DoesNotContain(Convert.ToBase64String(styleBytes));
    }

    // The embedded template is the file's own bytes, byte for byte.
    [Test]
    public async Task DocxTemplateBytesAreEmbedded()
    {
        var templateBytes = GeneratorDriver.BuildDocxBytes("Hello {{ Name }}");
        var result = Run(new TemplateFile("Report.parchment.docx", templateBytes));

        await Assert.That(Codes(result)).IsEmpty();
        await Assert.That(Generated(result)).Contains(Convert.ToBase64String(templateBytes));
    }
}
