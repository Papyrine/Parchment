// ReSharper disable PartialTypeWithSinglePart

public partial class LargeModelTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required IReadOnlyList<string> Items { get; init; }
    }

    // Fluid counts a step for every statement it executes, so the number of steps a render takes
    // grows with the model rather than with the template. A step limit therefore capped how much
    // data a template could be handed - and Fluid reports running out of steps as "The maximum level
    // of recursion has been reached. Your script must have a cyclic include statement", which sends
    // the reader looking for an include the template does not have.
    [Test]
    public async Task RenderIsNotCappedByTheSizeOfTheModel()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("{% for item in Items %}{{ item }}{% endfor %}", styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new Doc
            {
                Items = Enumerable.Repeat("x", 20_000).ToList()
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(doc.MainDocumentPart!.Document!.Body!.InnerText.Length).IsEqualTo(20_000);
    }
}
