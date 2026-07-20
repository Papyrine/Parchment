public class DeterminismTests
{
    // The double-render tests in this class compare renders taken milliseconds apart, which left a
    // hole: zip entry timestamps have 2-second resolution, entries cloned from the registration
    // snapshot keep their stamps, and a part ADDED during the render (settings, numbering, images)
    // was stamped with the wall clock. Two renders straddling a 2-second quantum differed in
    // exactly those bytes — a rare flake here, and a standing violation of the byte-identical
    // guarantee for any consumer rendering the same input twice on different days. Every entry is
    // now pinned to ZipTimestamps.StableDate, which these two tests assert directly since a
    // sleep-across-the-quantum repro is too slow for the suite.
    [Test]
    public async Task DocxRenderPinsZipEntryTimestamps()
    {
        using var template = DocxTemplateBuilder.Build("{{ Number }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>("docx-timestamps", template);

        using var stream = new MemoryStream();
        await store.Render("docx-timestamps", SampleData.Invoice(), stream);
        stream.Position = 0;

        await AssertPinnedTimestamps(stream);
    }

    [Test]
    public async Task MarkdownRenderPinsZipEntryTimestamps()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("markdown-timestamps", "# {{ Number }}", styleSource);

        using var stream = new MemoryStream();
        await store.Render("markdown-timestamps", SampleData.Invoice(), stream);
        stream.Position = 0;

        await AssertPinnedTimestamps(stream);
    }

    static async Task AssertPinnedTimestamps(MemoryStream stream)
    {
        using var archive = new System.IO.Compression.ZipArchive(stream);
        await Assert.That(archive.Entries.Count).IsGreaterThan(0);
        foreach (var entry in archive.Entries)
        {
            await Assert.That(entry.LastWriteTime.DateTime).IsEqualTo(ZipTimestamps.StableDate);
        }
    }

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
        // must be deterministic across renders, including the numbering instances a list creates
        // (OpenXmlHtml 1.0.6+ pins the numbering part's relationship id).
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
                Body = "<p>Hello <strong>world</strong> and <em>more</em></p><ul><li>a</li><li>b</li></ul>"
            },
            stream);
        return stream.ToArray();
    }

    [Test]
    public async Task EditableCollectionRenderIsByteIdentical()
    {
        // Repeating sections stamp a container sdt id, per-item ids and a perm-range id — all must be
        // deterministic across renders.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for b in Budgets %}

            Year: {{ b.Year }} Amount: {{ b.Amount }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableCollectionTests.BudgetPlan>("collection-determinism", template);

        var first = await RenderCollection(store);
        var second = await RenderCollection(store);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> RenderCollection(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(
            "collection-determinism",
            new EditableCollectionTests.BudgetPlan
            {
                Title = "T",
                Budgets =
                [
                    new()
                    {
                        Year = "2026",
                        Amount = 10m
                    },
                    new()
                    {
                        Year = "2027",
                        Amount = 20m
                    }
                ]
            },
            stream);
        return stream.ToArray();
    }
}
