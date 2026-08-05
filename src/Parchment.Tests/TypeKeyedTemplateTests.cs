// Registering and rendering without naming the template. The name is the part a caller most easily
// gets wrong — a string repeated at registration and at every render that the compiler cannot check
// — and where a model has one template the type already says which it is.
public class TypeKeyedTemplateTests
{
    public class Invoice
    {
        public required string Number { get; init; }
    }

    public class Receipt
    {
        public required string Number { get; init; }
    }

    static TemplateStore Store()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("# {{ Number }}");
        return store;
    }

    [Test]
    public async Task RegistersUnderTheModelsName()
    {
        var store = Store();

        // The name the generator's TemplateName emits, so a template registered either way is
        // reachable by the other.
        using var stream = new MemoryStream();
        await store.Render(nameof(Invoice), new Invoice { Number = "A-1" }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task RendersWithoutNamingTheTemplate()
    {
        var store = Store();

        using var stream = new MemoryStream();
        await store.Render(new Invoice { Number = "A-1" }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task RendersToFileWithoutNamingTheTemplate()
    {
        var store = Store();
        var path = Path.Combine(Path.GetTempPath(), $"parchment-{Guid.NewGuid():N}.docx");
        try
        {
            await store.RenderToFile(new Invoice { Number = "A-1" }, path);

            await Assert.That(new FileInfo(path).Length).IsGreaterThan(0);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public async Task ExplicitNameStillWins()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("chosen", "# {{ Number }}");

        using var stream = new MemoryStream();
        await store.Render(new Invoice { Number = "A-1" }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task ModelWithNoTemplateSaysSo()
    {
        var store = Store();

        using var stream = new MemoryStream();
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () => await store.Render(new Receipt { Number = "B-2" }, stream));

        await Assert.That(exception!.Message).Contains("No template is registered for Receipt");
    }

    // Two templates for one model leaves the type saying nothing about which to use. Rendering
    // through whichever was found first would depend on registration order, so the name is asked
    // for instead.
    [Test]
    public async Task ModelWithSeveralTemplatesAsksForTheName()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("summary", "# {{ Number }}");
        store.RegisterMarkdownTemplate<Invoice>("detail", "## {{ Number }}");

        using var stream = new MemoryStream();
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () => await store.Render(new Invoice { Number = "A-1" }, stream));

        await Assert.That(exception!.Message).Contains("2 templates are registered for Invoice");
        await Assert.That(exception!.Message).Contains("detail, summary");
    }
}
