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

    [Test]
    public async Task HtmlEditableFieldRenderIsByteIdentical()
    {
        // Editable HTML introduces a block sdt id, a perm-range id, and document protection — all
        // must be deterministic across renders. (Lists are intentionally excluded: when a template
        // has no numbering part, OpenXmlHtml adds one with a non-deterministic relationship id —
        // a pre-existing gap in that library that also affects read-only [Html], orthogonal to the
        // editable wrapper.)
        using var template = DocxTemplateBuilder.Build("{{ Body }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableFieldTests.EditableArticle>("html-editable-determinism", template);

        var first = await RenderHtmlEditable(store);
        var second = await RenderHtmlEditable(store);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> RenderHtmlEditable(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(
            "html-editable-determinism",
            new EditableFieldTests.EditableArticle
            {
                Title = "T",
                Body = "<p>Hello <strong>world</strong> and <em>more</em></p><p>Second paragraph.</p>"
            },
            stream);
        return stream.ToArray();
    }
}
