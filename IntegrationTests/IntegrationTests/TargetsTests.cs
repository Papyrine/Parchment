using System.Diagnostics;
using System.Runtime.CompilerServices;

/// <summary>
/// Cover for the MSBuild targets the package ships — the item types that declare a template, and
/// the warning for one declared under two item types.
/// </summary>
/// <remarks>
/// These only exist inside a build, so the test runs one: a fixture project is compiled against the
/// packed Parchment and its output is read back. Nothing about them is observable from a running
/// process, and a unit test over the targets file would only assert that XML says what it says.
/// </remarks>
public class TargetsTests
{
    static string FixtureProject([CallerFilePath] string file = "") =>
        Path.Combine(
            Path.GetDirectoryName(Path.GetDirectoryName(file))!,
            "TargetsFixture",
            "TargetsFixture.csproj");

    static string BuildOutput { get; set; } = "";

    [Before(Class)]
    public static async Task BuildFixture()
    {
        var project = FixtureProject();
        // Fresh every run: MSBuild skips targets for an up-to-date project, and a skipped check
        // reports nothing, which would read as a passing test.
        foreach (var directory in new[] { "obj", "bin" })
        {
            var path = Path.Combine(Path.GetDirectoryName(project)!, directory);
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }

        using var process = new Process
        {
            StartInfo = new()
            {
                FileName = "dotnet",
                Arguments = $"build \"{project}\" --nologo",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            }
        };
        process.Start();
        var output = await process.StandardOutput.ReadToEndAsync();
        var error = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        BuildOutput = output + error;
    }

    // A warning must not become a failure: the misconfiguration PARCH100 reports is one the command
    // line builds through, which is the whole reason it needs reporting.
    [Test]
    public async Task FixtureBuilds() =>
        await Assert.That(BuildOutput).Contains("Build succeeded");

    [Test]
    public async Task WarnsForTheTemplateCarryingTwoItemTypes()
    {
        await Assert.That(BuildOutput).Contains("PARCH100");
        await Assert.That(BuildOutput).Contains("dual.md");
    }

    // The embedded item type is the one at risk of a false positive: it puts the template in
    // AdditionalFiles and a staged copy in EmbeddedResource, so a check comparing item types
    // without comparing paths would flag it.
    [Test]
    public async Task DoesNotWarnForTemplatesDeclaredWithTheItemTypes()
    {
        await Assert.That(BuildOutput).DoesNotContain("copied.md");
        await Assert.That(BuildOutput).DoesNotContain("embedded.md");
    }

    // Proves the item types actually reach the generator. Without this the suite would pass on a
    // targets file that declared nothing at all.
    //
    // Matched as "error PARCH004" rather than the bare code, because PARCH100's own text names
    // PARCH004 as the thing it prevents — the loose match found the cure and called it the disease.
    [Test]
    public async Task EveryTemplateIsVisibleToTheGenerator() =>
        await Assert.That(BuildOutput).DoesNotContain("error PARCH004");
}
