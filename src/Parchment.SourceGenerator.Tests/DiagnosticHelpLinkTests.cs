// The help link is what an IDE turns into "why is this here?", so a diagnostic that points at a
// heading the readme no longer has is worse than one with no link at all — it looks answered.
// Renaming a PARCH heading breaks these, which is the point.
public class DiagnosticHelpLinkTests
{
    const string Prefix = "https://github.com/Papyrine/Parchment#";

    static IReadOnlyList<DiagnosticDescriptor> Descriptors =>
        typeof(Diagnostics)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(_ => _.FieldType == typeof(DiagnosticDescriptor))
            .Select(_ => (DiagnosticDescriptor)_.GetValue(null)!)
            .ToList();

    [Test]
    public async Task EveryDiagnosticLinksToTheReadme()
    {
        foreach (var descriptor in Descriptors)
        {
            await Assert.That(descriptor.HelpLinkUri)
                .StartsWith(Prefix)
                .Because($"{descriptor.Id} should link to its readme section");
        }
    }

    [Test]
    public async Task EveryHelpLinkResolvesToAHeadingThatExists()
    {
        var anchors = ReadmeAnchors();
        foreach (var descriptor in Descriptors)
        {
            var anchor = descriptor.HelpLinkUri[Prefix.Length..];
            await Assert.That(anchors)
                .Contains(anchor)
                .Because($"{descriptor.Id} links to '#{anchor}', which no readme heading produces");
        }
    }

    // GitHub builds a heading anchor by lowercasing, dropping everything that is not alphanumeric,
    // a space or a hyphen, then turning spaces into hyphens.
    static List<string> ReadmeAnchors()
    {
        var path = Path.Combine(SolutionDirectory(), "..", "readme.md");
        var anchors = new List<string>();
        foreach (var line in File.ReadAllLines(path))
        {
            if (!line.StartsWith("### ", StringComparison.Ordinal))
            {
                continue;
            }

            var heading = line[4..].ToLowerInvariant();
            var builder = new StringBuilder();
            foreach (var character in heading)
            {
                if (char.IsLetterOrDigit(character) ||
                    character == '-')
                {
                    builder.Append(character);
                }
                else if (character == ' ')
                {
                    builder.Append('-');
                }
            }

            anchors.Add(builder.ToString());
        }

        return anchors;
    }

    static string SolutionDirectory([CallerFilePath] string file = "") =>
        Path.GetDirectoryName(Path.GetDirectoryName(file))!;
}
