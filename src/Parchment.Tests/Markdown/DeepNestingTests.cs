public class DeepNestingTests
{
    public class EmptyModel;

    [Test]
    public async Task QuoteContainingListContainingNestedList()
    {
        // Exercises ContainerState push/pop discipline: quote → list → nested list → list-item
        // text. If the stack is mismatched at any pop, paragraphs end up with the wrong style or
        // the wrong NumberingProperties.
        const string md =
            """
            > Top quote

            > - outer one
            >   - inner a
            >   - inner b
            > - outer two

            Tail
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(md, styleSource);

        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task ListContainingQuoteWithEmphasis()
    {
        // List → paragraph → quote → emphasized inline. The renderer has to push a list
        // ContainerState, push a quote ContainerState, accumulate emphasized runs, then unwind
        // both pops in the correct order.
        const string md =
            """
            - first item

              > *nested quote with **bold*** text

            - second item
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(md, styleSource);

        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task SixHeadingLevelsAllRenderWithStyles()
    {
        const string md =
            """
            # H1

            ## H2

            ### H3

            #### H4

            ##### H5

            ###### H6
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(md, styleSource);

        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var styles = doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Select(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
            .Where(_ => _ != null && _.StartsWith("Heading"))
            .Select(_ => _!)
            .ToList();
        await Assert.That(styles).IsEquivalentTo([
            "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6"
        ]);
    }
}
