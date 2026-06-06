/// <summary>
/// Regenerates an <c>input.png</c> alongside every <c>scenarios/*/input.docx</c> by running the
/// template docx through Morph's SkiaSharp-backed renderer. Marked <c>[Explicit]</c> because it
/// only needs to run when a scenario's input or the renderer changes — the png is a docs asset,
/// not a test assertion.
/// </summary>
public class ScenarioInputRenderer
{
    [Test, Explicit]
    public async Task RenderAllInputDocxesToPng()
    {
        var scenariosDir = Path.GetFullPath(Path.Combine(
            Path.GetDirectoryName(SourcePath())!,
            ".."));

        var inputs = Directory.GetFiles(scenariosDir, "input.docx", SearchOption.AllDirectories);
        await Assert.That(inputs.Length).IsGreaterThan(0);

        var converter = new SkiaDocumentConverter();
        var options = new ImageExportOptions();

        foreach (var docxPath in inputs)
        {
            await using var stream = File.OpenRead(docxPath);
            var pages = converter.ConvertToImageData(stream, options);
            await Assert.That(pages.Count).IsGreaterThan(0);

            var pngPath = Path.Combine(Path.GetDirectoryName(docxPath)!, "input.png");
            await File.WriteAllBytesAsync(pngPath, pages[0]);
        }
    }

    static string SourcePath([CallerFilePath] string path = "") => path;
}
