/// <summary>
/// Cover for the MSBuild targets the package ships — the globs that discover a template by name,
/// the item type that declares one they cannot reach, and the warning for a template declared
/// under two item types.
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

    // The items the build worked from. Item state is not observable in build output, and printing
    // it there would collide with the assertions that read that output for file names.
    static string NoneItems { get; set; } = "";

    static string AdditionalFileItems { get; set; } = "";

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

        BuildOutput = await Dotnet($"build \"{project}\" --nologo");
        // After the build, so the restore it needs has already run.
        NoneItems = await Dotnet($"msbuild \"{project}\" -getItem:None --nologo");
        AdditionalFileItems = await Dotnet($"msbuild \"{project}\" -getItem:AdditionalFiles --nologo");
    }

    static async Task<string> Dotnet(string arguments)
    {
        using var process = new Process
        {
            StartInfo = new()
            {
                FileName = "dotnet",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            }
        };
        process.Start();
        var output = await process.StandardOutput.ReadToEndAsync();
        var error = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return output + error;
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
        await Assert.That(BuildOutput).Contains("DualModel.parchment.md");
    }

    [Test]
    public async Task DoesNotWarnForTemplatesDeclaredWithTheItemType() =>
        await Assert.That(BuildOutput).DoesNotContain("CopiedModel.parchment.md");

    // Proves the item types actually reach the generator. Without this the suite would pass on a
    // targets file that declared nothing at all.
    //
    // Matched as "error PARCH004" rather than the bare code, because PARCH100's own text names
    // PARCH004 as the thing it prevents — the loose match found the cure and called it the disease.
    [Test]
    public async Task EveryTemplateIsVisibleToTheGenerator() =>
        await Assert.That(BuildOutput).DoesNotContain("error PARCH004");

    // The SDK's default glob sweeps every template into None as well, and the package leaves that
    // identity standing: None never reaches the compiler, so it cannot shadow the AdditionalFiles
    // entry, and it is what keeps a template listed in a solution explorer — an IDE lists project
    // files through item types like None, and a template stripped of it disappears from the tree.
    [Test]
    public async Task TemplatesKeepTheirNoneIdentity()
    {
        await Assert.That(NoneItems).Contains("CopiedModel.parchment.md");
        await Assert.That(NoneItems).Contains("DualModel.parchment.md");
        await Assert.That(NoneItems).Contains("DiscoveredModel.parchment.md");
    }

    // The marker earns its keep here: DiscoveredModel.parchment.md carries no csproj entry, so the
    // globs are the only thing that can put it in front of the generator. Its model would fail the
    // build with PARCH004 if they did not — which FixtureBuilds and EveryTemplateIsVisibleToTheGenerator
    // already assert — and this names the file the assertion depends on.
    [Test]
    public async Task DiscoversAnUndeclaredTemplate() =>
        await Assert.That(AdditionalFileItems).Contains("DiscoveredModel.parchment.md");

    // An unmarked markdown file is not swept in, which is the whole reason globbing is safe.
    [Test]
    public async Task DoesNotDiscoverUnmarkedMarkdown() =>
        await Assert.That(AdditionalFileItems).DoesNotContain("notes.md");
}
