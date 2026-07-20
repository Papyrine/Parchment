using Fluid;
using Fluid.Values;

public class AddFilterTests
{
    public class Doc
    {
        public required string Body { get; init; }
    }

    // AddFilter mutates process-wide static state, so the name is unique to this class and is
    // registered once rather than once per test.
    static AddFilterTests() =>
        TemplateStore.AddFilter("add_filter_test_shout", Shout);

    static ValueTask<FluidValue> Shout(FluidValue input, FilterArguments arguments, TemplateContext context) =>
        new(new Fluid.Values.StringValue(input.ToStringValue().ToUpperInvariant()));

    // The two flows render through separate TemplateOptions, and Fluid exposes Filters as get-only,
    // so the two cannot be pointed at one collection. Registering against the docx set alone left
    // the filter invisible here — and Fluid passes an unknown filter's input straight through, so
    // the value rendered unfiltered with nothing reported.
    [Test]
    public async Task RegisteredFilterAppliesInTheMarkdownFlow()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("md", "{{ Body | add_filter_test_shout }}", styleSource);

        await Assert.That(await RenderText(store, "md")).IsEqualTo("QUIET");
    }

    [Test]
    public async Task RegisteredFilterAppliesInTheDocxFlow()
    {
        using var template = DocxTemplateBuilder.Build("{{ Body | add_filter_test_shout }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("docx", template);

        await Assert.That(await RenderText(store, "docx")).IsEqualTo("QUIET");
    }

    static async Task<string> RenderText(TemplateStore store, string name)
    {
        using var stream = new MemoryStream();
        await store.Render(
            name,
            new Doc
            {
                Body = "quiet"
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
