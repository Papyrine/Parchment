using DocumentFormat.OpenXml.Validation;

/// <summary>
/// Holds every checked-in .docx snapshot to the schema Word enforces.
/// </summary>
/// <remarks>
/// The snapshots are the recorded output of every flow this suite covers, which makes sweeping them
/// the cheapest way to hold all of it to the schema at once - the render calls themselves are spread
/// across forty files with no shared harness to hang an assertion on.
/// <para>
/// Nothing else catches this. Morph reaches for elements through the OpenXML object model, which
/// finds a child wherever it sits inside its parent, so a document Word would refuse still paginates
/// and renders, and the page snapshots agree with it. A w:pPr written after the runs rather than
/// before them cost a table of contents its dot leader in Word while every test here stayed green.
/// </para>
/// </remarks>
public class SchemaValidationTests
{
    [Test]
    [MethodDataSource(nameof(Snapshots))]
    public async Task SnapshotMatchesTheSchema(string snapshot)
    {
        using var doc = WordprocessingDocument.Open(Path.Combine(ProjectFiles.ProjectDirectory, snapshot), false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);

        await Assert.That(validator.Validate(doc).Select(_ => $"{_.Description} @ {_.Path?.XPath}")).IsEmpty();
    }

    // A sweep that stops matching leaves nothing to run and nothing to report, which reads exactly
    // like a clean one.
    [Test]
    public async Task SnapshotsAreFound() =>
        await Assert.That(Snapshots()).IsNotEmpty();

    public static IEnumerable<string> Snapshots() =>
        Directory.EnumerateFiles(ProjectFiles.ProjectDirectory, "*.verified.docx", SearchOption.AllDirectories)
            .Select(_ => Path.GetRelativePath(ProjectFiles.ProjectDirectory, _).Replace(Path.DirectorySeparatorChar, '/'))
            // Build output carries copies of the snapshots; they are the same files twice over.
            .Where(_ => !_.StartsWith("bin/", StringComparison.Ordinal) &&
                        !_.StartsWith("obj/", StringComparison.Ordinal))
            .Order(StringComparer.Ordinal);
}
