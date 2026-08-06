// Covers TemplateStore.Materialize — the source-generator cold-start path — which nothing reaches
// today. RegistrationBenchmarks.RegisterViaSourceGeneratorPath reads as though it does, but it
// pre-seeds the Fluid accessors and then calls the public RegisterDocxTemplate overload; Materialize
// is only entered through Render -> Resolve -> GeneratedTemplateDefinitions.TryGet. ParchmentModel's
// module initializer registers a docx definition for Invoice and a markdown one for ReportContext,
// so a store that has never seen the type materializes the real generated definition on first
// render.
//
// Materialize is private and not separately callable, so each flow is measured as a pair: first
// render on a cold store against a render on a warm one. The difference is registration. On the
// docx side that currently round-trips the definition's byte[] into a MemoryStream, back out to a
// byte[], and into a MemoryStream again before any work starts.
[Config(typeof(BenchmarkConfig))]
public class TemplateMaterializeBenchmarks
{
    Invoice invoice = null!;
    ReportContext report = null!;
    TemplateStore warmDocxStore = null!;
    TemplateStore warmMarkdownStore = null!;

    [GlobalSetup]
    public async Task Setup()
    {
        invoice = SampleData.Invoice();
        report = SampleData.Report();

        // One render materializes the definition into the store, leaving the warm benchmarks
        // measuring the render alone.
        warmDocxStore = new();
        using var docxWarmup = new MemoryStream();
        await warmDocxStore.Render(invoice, docxWarmup);

        warmMarkdownStore = new();
        using var markdownWarmup = new MemoryStream();
        await warmMarkdownStore.Render(report, markdownWarmup);
    }

    [Benchmark]
    public async Task DocxColdStore()
    {
        var store = new TemplateStore();
        using var output = new MemoryStream();
        await store.Render(invoice, output);
    }

    [Benchmark]
    public async Task DocxWarmStore()
    {
        using var output = new MemoryStream();
        await warmDocxStore.Render(invoice, output);
    }

    [Benchmark]
    public async Task MarkdownColdStore()
    {
        var store = new TemplateStore();
        using var output = new MemoryStream();
        await store.Render(report, output);
    }

    [Benchmark]
    public async Task MarkdownWarmStore()
    {
        using var output = new MemoryStream();
        await warmMarkdownStore.Render(report, output);
    }
}
