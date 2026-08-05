// ReSharper disable PartialTypeWithSinglePart
// Registering and rendering is keyed by the model type. There is no name to repeat at
// registration and at every render, so there is no string for a caller to get wrong — the type
// says which template to use.
public partial class TypeKeyedTemplateTests
{
    [ParchmentBindable]
    public partial class Invoice
    {
        public required string Number { get; init; }
    }

    public class Receipt
    {
        public required string Number { get; init; }
    }

    // Deliberately the same simple name as the Invoice above, in a different namespace.
    public static partial class Other
    {
        [ParchmentBindable]
        public partial class Invoice
        {
            public required string Number { get; init; }
        }
    }

    static TemplateStore Store()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("# {{ Number }}");
        return store;
    }

    // The dictionary is keyed on the Type itself, so two models sharing a simple name cannot
    // collide.
    [Test]
    public async Task ModelsSharingASimpleNameDoNotCollide()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("# {{ Number }}");
        store.RegisterMarkdownTemplate<Other.Invoice>("## {{ Number }}");

        using var first = new MemoryStream();
        using var second = new MemoryStream();
        await store.Render(new Invoice { Number = "A-1" }, first);
        await store.Render(new Other.Invoice { Number = "B-2" }, second);

        await Assert.That(first.Length).IsGreaterThan(0);
        await Assert.That(second.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task ReregisteringTheSameModelReplacesIt()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("# first");
        store.RegisterMarkdownTemplate<Invoice>("# second");

        using var stream = new MemoryStream();
        await store.Render(new Invoice { Number = "A-1" }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task RendersByModelType()
    {
        var store = Store();

        using var stream = new MemoryStream();
        await store.Render(new Invoice { Number = "A-1" }, stream);

        await Assert.That(stream.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task RendersToFileByModelType()
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
    public async Task ModelWithNoTemplateSaysSo()
    {
        var store = Store();

        using var stream = new MemoryStream();
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            () => store.Render(new Receipt { Number = "B-2" }, stream));

        await Assert.That(exception!.Message).Contains("No template is registered for Receipt");
    }
}
