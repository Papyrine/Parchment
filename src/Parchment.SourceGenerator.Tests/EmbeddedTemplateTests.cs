/// <summary>
/// How template bytes travel in the generated source: as the argument of an attribute, which the
/// compiler stores in the blob heap, not as a string the code loads, which it stores in the
/// user-string heap that every literal of the assembly shares. These compile the generated output
/// against the real Parchment assembly — the limit was the compiler's, and only showed at emit.
/// </summary>
public class EmbeddedTemplateTests
{
    static string Models(int count)
    {
        var builder = new StringBuilder(
            """
            using Parchment;

            namespace Sample;

            """);
        for (var index = 0; index < count; index++)
        {
            builder.AppendLine($"[ParchmentModel] public partial class M{index} {{ public string Name {{ get; set; }} = \"\"; }}");
        }

        return builder.ToString();
    }

    // A docx padded with an incompressible image part, so its size is what was asked for.
    static byte[] DocxOfSize(Random random, int payloadBytes, string paragraph = "Hello {{ Name }}")
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body(GeneratorDriver.BuildParagraph(paragraph)));
            var settings = mainPart.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new(new RemovePersonalInformation(), new RemoveDateAndTime());
            var payload = new byte[payloadBytes];
            random.NextBytes(payload);
            mainPart.AddImagePart(ImagePartType.Png).FeedData(new MemoryStream(payload));
        }

        return stream.ToArray();
    }

    static string Describe(CompiledGenerator.Result result) =>
        string.Join("\n", result.Errors.Select(_ => $"{_.Id}: {_.GetMessage()}"));

    static byte[] Embedded(System.Reflection.Assembly assembly, string model)
    {
        var type = assembly.GetType($"Sample.{model}+ParchmentGenerated_Sample_{model}", throwOnError: true)!;
        var attribute = (Parchment.Generated.EmbeddedBytesAttribute) type
            .GetCustomAttributes(typeof(Parchment.Generated.EmbeddedBytesAttribute), false)
            .Single();
        return Convert.FromBase64String(attribute.Base64);
    }

    static async Task<string> Render(System.Reflection.Assembly assembly, string model)
    {
        dynamic instance = Activator.CreateInstance(assembly.GetType($"Sample.{model}", throwOnError: true)!)!;
        instance.Name = "World";
        var store = new Parchment.TemplateStore();
        using var stream = new MemoryStream();
        await (Task) store.Render(instance, stream);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }

    // Nine templates of 800 KB: 7.2 MB in all, past the 6.2 MB at which base64 string literals
    // filled the user-string heap and the build stopped with CS8103.
    [Test]
    public async Task TemplatesTotallingMoreThanTheStringHeapCompile()
    {
        var random = new Random(42);
        var files = Enumerable.Range(0, 9)
            .Select(_ => new TemplateFile($"M{_}.parchment.docx", DocxOfSize(random, 800 * 1024)))
            .ToArray();
        var total = files.Sum(_ => (long) _.Bytes.Length);

        var result = CompiledGenerator.Run(Models(9), files);

        await Assert.That(Describe(result)).IsEqualTo("");
        // Base64 as UTF-8 is four thirds of the template; as a loaded string it was eight thirds.
        await Assert.That((long) result.Assembly.Length).IsLessThan(total * 3 / 2);

        var assembly = System.Reflection.Assembly.Load(result.Assembly);
        await Assert.That(Embedded(assembly, "M8").SequenceEqual(files[8].Bytes)).IsTrue();
        await Assert.That(await Render(assembly, "M8")).Contains("Hello World");
    }

    // One template is enough: 20 MB is more than the heap could ever hold.
    [Test]
    public async Task SingleTemplateLargerThanTheStringHeapCompiles()
    {
        var docx = DocxOfSize(new(1), 20 * 1024 * 1024);

        var result = CompiledGenerator.Run(Models(1), [new("M0.parchment.docx", docx)]);

        await Assert.That(Describe(result)).IsEqualTo("");
        var assembly = System.Reflection.Assembly.Load(result.Assembly);
        await Assert.That(Embedded(assembly, "M0").SequenceEqual(docx)).IsTrue();
    }

    // Both flows, through the real runtime.
    [Test]
    public async Task BothFlowsRenderFromTheAttribute()
    {
        var random = new Random(7);
        var docx = DocxOfSize(random, 1001);
        var style = DocxOfSize(random, 3001, "styled");
        TemplateFile[] files =
        [
            new("M0.parchment.docx", docx),
            new("M1.parchment.md", Encoding.UTF8.GetBytes("Greetings {{ Name }}")),
            new("M1.parchment.dotx", style),
            new("M2.parchment.md", Encoding.UTF8.GetBytes("Plain {{ Name }}"))
        ];

        var result = CompiledGenerator.Run(Models(3), files);

        await Assert.That(Describe(result)).IsEqualTo("");
        var assembly = System.Reflection.Assembly.Load(result.Assembly);
        await Assert.That(Embedded(assembly, "M0").SequenceEqual(docx)).IsTrue();
        await Assert.That(Embedded(assembly, "M1").SequenceEqual(style)).IsTrue();
        await Assert.That(await Render(assembly, "M0")).Contains("Hello World");
        await Assert.That(await Render(assembly, "M1")).Contains("Greetings World");
        await Assert.That(await Render(assembly, "M2")).Contains("Plain World");
        // A markdown model with no style document carries no bytes at all.
        await Assert.That(result.Generated.Split("EmbeddedBytes(\"").Length - 1).IsEqualTo(2);
    }

    // Every model's generated file carries the shared style document, and the assembly carries it
    // once: identical attribute values are one entry in the blob heap.
    [Test]
    public async Task SharedStyleDocumentIsStoredOnce()
    {
        var style = DocxOfSize(new(3), 150 * 1024, "styled");
        var files = new List<TemplateFile>
        {
            new("parchment.dotx", style)
        };
        for (var index = 0; index < 20; index++)
        {
            files.Add(new($"M{index}.parchment.md", Encoding.UTF8.GetBytes("# {{ Name }}")));
        }

        var result = CompiledGenerator.Run(Models(20), files.ToArray());

        await Assert.That(Describe(result)).IsEqualTo("");
        await Assert.That(result.Generated.Split("EmbeddedBytes(\"").Length - 1).IsEqualTo(20);
        await Assert.That(result.Assembly.Length).IsLessThan(style.Length * 2);
    }
}
