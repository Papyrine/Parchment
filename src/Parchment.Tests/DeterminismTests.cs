public partial class DeterminismTests
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
        store.RegisterDocxTemplate<Invoice>(template);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);
        stream.Position = 0;

        await AssertPinnedTimestamps(stream);
    }

    [Test]
    public async Task MarkdownRenderPinsZipEntryTimestamps()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>("# {{ Number }}", styleSource);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);
        stream.Position = 0;

        await AssertPinnedTimestamps(stream);
    }

    static async Task AssertPinnedTimestamps(MemoryStream stream)
    {
        await using var archive = new System.IO.Compression.ZipArchive(stream);
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
        store.RegisterDocxTemplate<Invoice>(template);

        var first = await Render(store, SampleData.Invoice());
        var second = await Render(store, SampleData.Invoice());

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> Render(TemplateStore store, Invoice model)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
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
        store.RegisterDocxTemplate<EditableFieldTests.EditableOrder>(template);

        var first = await RenderEditable(store);
        var second = await RenderEditable(store);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> RenderEditable(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(EditableFieldTests.NewOrder(), stream);
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
        store.RegisterDocxTemplate<EditableFieldTests.EditableArticle>(template);

        var first = await RenderHtmlEditable(store);
        var second = await RenderHtmlEditable(store);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> RenderHtmlEditable(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(
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
        store.RegisterDocxTemplate<EditableCollectionTests.BudgetPlan>(template);

        var first = await RenderCollection(store);
        var second = await RenderCollection(store);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static async Task<byte[]> RenderCollection(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(
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

    // Everything below covers one family: a part or relationship added WITHOUT an explicit id is
    // given a random one by the Open XML SDK ("R" + 16 hex digits), which lands in the .rels file
    // and in the XML that references it. Every such add now goes through RelationshipIds.
    [ParchmentBindable]
    public partial class Linked
    {
        public string Name { get; set; } = "Alice";
        public string[] Tags { get; set; } = [];
        public OpenXmlToken? Logo { get; set; }
    }

    [ParchmentBindable]
    public partial class Editable
    {
        [EditableField]
        public string? PurchaseOrder { get; set; }
    }

    // No styles, settings, numbering or property parts, so each of them is one Parchment has to add.
    static MemoryStream BareDocx(string text, Action<MainDocumentPart>? customize = null)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new(
                new Body(
                    new Paragraph(new Run(new Text(text))),
                    new SectionProperties(
                        new PageSize
                        {
                            Width = 11906,
                            Height = 16838
                        })));
            customize?.Invoke(mainPart);
        }

        stream.Position = 0;
        return stream;
    }

    static async Task<byte[]> RenderBytes<T>(TemplateStore store, T model, WordDocumentProperties? properties = null)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream, properties);
        return stream.ToArray();
    }

    [Test]
    public async Task MarkdownLinkRenderIsByteIdentical()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Linked>(
            "See [site](https://example.com), [other](https://example.org) and [site again](https://example.com) {{ Name }}");

        var first = await RenderBytes(store, new Linked());
        var second = await RenderBytes(store, new Linked());

        await Assert.That(first).IsEquivalentTo(second);

        // Each target gets its own relationship, and a repeated target shares one.
        using var document = WordprocessingDocument.Open(new MemoryStream(first), false);
        var mainPart = document.MainDocumentPart!;
        var ids = mainPart.Document!.Descendants<Hyperlink>().Select(_ => _.Id!.Value!).ToList();
        await Assert.That(ids.Count).IsEqualTo(3);
        await Assert.That(ids[0]).IsNotEqualTo(ids[1]);
        await Assert.That(ids[2]).IsEqualTo(ids[0]);
        await Assert.That(mainPart.HyperlinkRelationships.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task MarkdownAutolinkRenderIsByteIdentical()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Linked>("See <https://example.com> or <someone@example.com> {{ Name }}");

        var first = await RenderBytes(store, new Linked());
        var second = await RenderBytes(store, new Linked());

        await Assert.That(first).IsEquivalentTo(second);
    }

    // A url with a bound tail, bare and in angle brackets. The bare form is a LinkInline built by
    // BoundAutoLinkParser and the angle form an AutolinkInline, so the two take different renderers.
    [Test]
    public async Task MarkdownUrlWithBoundTailRenderIsByteIdentical()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Linked>(
            "See https://example.com/orders/{{ Name }} or <https://example.org/orders/{{ Name }}> now");

        var first = await RenderBytes(store, WithEscapedName());
        var second = await RenderBytes(store, WithEscapedName());

        await Assert.That(first).IsEquivalentTo(second);

        using var document = WordprocessingDocument.Open(new MemoryStream(first), false);
        await Assert.That(document.MainDocumentPart!.HyperlinkRelationships.Select(_ => _.Uri.OriginalString))
            .IsEquivalentTo(["https://example.com/orders/A-1_x.y", "https://example.org/orders/A-1_x.y"]);
    }

    // A value the markdown encoder has to escape, so the unescaped url is what reaches the relationship.
    static Linked WithEscapedName() =>
        new()
        {
            Name = "A-1_x.y"
        };

    [Test]
    public async Task MarkdownListOnBlankStyleSourceRenderIsByteIdentical()
    {
        // The built-in style source has no numbering part, so the list has to add one.
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Linked>("- a\n- b\n\n1. {{ Name }}\n");

        var first = await RenderBytes(store, new Linked());
        var second = await RenderBytes(store, new Linked());

        await Assert.That(first).IsEquivalentTo(second);
    }

    [Test]
    public async Task BulletListWithoutNumberingPartRenderIsByteIdentical()
    {
        using var template = BareDocx("{{ Tags | bullet_list }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Linked>(template);

        var first = await RenderBytes(store, WithTags());
        var second = await RenderBytes(store, WithTags());

        await Assert.That(first).IsEquivalentTo(second);
    }

    static Linked WithTags() =>
        new()
        {
            Tags = ["a", "b"]
        };

    [Test]
    public async Task DocumentPropertiesWithoutPropertyPartsRenderIsByteIdentical()
    {
        // Core, extended and custom: none of the three exists in the template.
        using var template = BareDocx("{{ Name }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Linked>(template);

        var first = await RenderBytes(store, new Linked(), Properties());
        var second = await RenderBytes(store, new Linked(), Properties());

        await Assert.That(first).IsEquivalentTo(second);
    }

    static WordDocumentProperties Properties() =>
        new()
        {
            Title = "Title",
            Company = "Company",
            Custom =
            {
                ["Reference"] = "x"
            }
        };

    [Test]
    public async Task ImagePartRenderIsByteIdentical()
    {
        using var template = BareDocx("{{ Logo }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Linked>(template);

        var first = await RenderBytes(store, WithLogo());
        var second = await RenderBytes(store, WithLogo());

        await Assert.That(first).IsEquivalentTo(second);

        using var document = WordprocessingDocument.Open(new MemoryStream(first), false);
        await Assert.That(document.MainDocumentPart!.ImageParts.Count()).IsEqualTo(2);
    }

    // Two image parts from one token, so the second id has to step past the first.
    static Linked WithLogo() =>
        new()
        {
            Tags = ["a", "b"],
            Logo = new(context =>
            {
                var first = context.AddImagePart(OnePixelPng(), "image/png");
                var second = context.AddImagePart(OnePixelPng(), "image/png");
                return [new Paragraph(new Run(new Text($"{first} {second}")))];
            })
        };

    static byte[] OnePixelPng() =>
        Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=");

    // The built-in style source is built once per process and cached, so a double render cannot see
    // a random id in it: both renders clone the same bytes. Two builds stand in for two processes.
    [Test]
    public async Task BlankStyleSourceHasStableRelationshipIds()
    {
        var first = BlankRelationshipIds();
        var second = BlankRelationshipIds();

        await Assert.That(first).IsEquivalentTo(second);
        await Assert.That(first).IsEquivalentTo(["rId1", "rIdStyles"]);
    }

    static string[] BlankRelationshipIds()
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(TemplateStore.BuildBlankDocx()), false);
        return
        [
            document.GetIdOfPart(document.MainDocumentPart!),
            ..document.MainDocumentPart!.Parts.Select(_ => _.RelationshipId)
        ];
    }

    // Protection is written at registration, so here it is the same template bytes registered twice
    // that have to agree, rather than one registration rendered twice.
    [Test]
    public async Task EditableWithoutSettingsPartIsByteIdenticalAcrossRegistrations()
    {
        byte[] template;
        using (var stream = BareDocx("{{ PurchaseOrder }}"))
        {
            template = stream.ToArray();
        }

        var first = await RenderRegistered(template);
        var second = await RenderRegistered(template);

        await Assert.That(first).IsEquivalentTo(second);
    }

    static Task<byte[]> RenderRegistered(byte[] template)
    {
        var store = new TemplateStore();
        using var stream = new MemoryStream(template);
        store.RegisterDocxTemplate<Editable>(stream);
        return RenderBytes(
            store,
            new Editable
            {
                PurchaseOrder = "1"
            });
    }

    // A template's own ids are arbitrary, so it may already be using the very ids Parchment would
    // otherwise pick. Each add has to step past them rather than throw on the duplicate.
    [Test]
    public async Task AddedPartsStepPastIdsTheTemplateAlreadyUses()
    {
        string[] taken = ["rIdSettings", "rIdNumbering", "rIdImage1", "rIdLink1"];
        using var template = BareDocx(
            "{{ Tags | bullet_list }}{{ Logo }}",
            mainPart =>
            {
                foreach (var id in taken)
                {
                    mainPart.AddHyperlinkRelationship(new($"https://example.com/{id}"), true, id);
                }
            });
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Linked>(template);

        var first = await RenderBytes(store, WithLogo());
        var second = await RenderBytes(store, WithLogo());

        await Assert.That(first).IsEquivalentTo(second);

        using var document = WordprocessingDocument.Open(new MemoryStream(first), false);
        var mainPart = document.MainDocumentPart!;
        await Assert.That(mainPart.HyperlinkRelationships.Select(_ => _.Id)).IsEquivalentTo(taken);
        await Assert.That(mainPart.NumberingDefinitionsPart).IsNotNull();
        await Assert.That(mainPart.DocumentSettingsPart).IsNotNull();
        await Assert.That(mainPart.ImageParts.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task MarkdownLinksStepPastIdsTheStyleSourceAlreadyUses()
    {
        using var styleSource = BareDocx(
            "",
            mainPart => mainPart.AddHyperlinkRelationship(new("https://example.com/own"), true, "rIdLink1"));
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Linked>("[site](https://example.com) {{ Name }}", styleSource);

        var first = await RenderBytes(store, new Linked());
        var second = await RenderBytes(store, new Linked());

        await Assert.That(first).IsEquivalentTo(second);

        using var document = WordprocessingDocument.Open(new MemoryStream(first), false);
        var mainPart = document.MainDocumentPart!;
        var id = mainPart.Document!.Descendants<Hyperlink>().Single().Id!.Value!;
        await Assert.That(mainPart.HyperlinkRelationships.Single(_ => _.Id == id).Uri.OriginalString)
            .IsEqualTo("https://example.com");
    }
}
