public class DeterminismTests
{
    [Test]
    public async Task DocxRenderIsByteIdentical()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Invoice {{ Number }}

            Customer: {{ Customer.Name }}

            Total: {{ Total }} {{ Currency }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>("determinism", template);

        var first = await Render(store, SampleData.Invoice());
        var second = await Render(store, SampleData.Invoice());

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> Render(TemplateStore store, Invoice model)
    {
        using var stream = new MemoryStream();
        await store.Render("determinism", model, stream);
        return stream.ToArray();
    }

    [Test]
    public async Task EditableFieldRenderIsByteIdentical()
    {
        // Editable fields introduce sdt ids, perm-range ids, and document protection — all must
        // be deterministic (sequential ids, passwordless enforcement, no timestamps).
        using var template = DocxTemplateBuilder.Build(
            """
            PO: {{ PurchaseOrder }}

            {{ Approved }} {{ Delivery }} {{ Status }} {{ Discount }}

            {{ Notes }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableFieldTests.EditableOrder>("editable-determinism", template);

        var first = await RenderEditable(store);
        var second = await RenderEditable(store);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> RenderEditable(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render("editable-determinism", EditableFieldTests.NewOrder(), stream);
        return stream.ToArray();
    }
}
