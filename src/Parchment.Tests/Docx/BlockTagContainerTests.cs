// The runner walks a block's content by sibling, from its open tag to its close tag, so the
// paragraphs holding one block's tags have to share a parent. These pin that registration says so
// instead of leaving the render to empty cells or delete what follows.
public class BlockTagContainerTests
{
    [Test]
    public async Task ForAndEndForInDifferentTableRowsIsRejected()
    {
        using var template = BuildTemplate(
            BuildParagraph("Before"),
            new Table(
                new TableRow(new TableCell(BuildParagraph("{% for line in Lines %}"))),
                new TableRow(new TableCell(BuildParagraph("{{ line.Description }}"))),
                new TableRow(new TableCell(BuildParagraph("{% endfor %}")))),
            BuildParagraph("After"));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Invoice>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("{% for line in Lines %}");
        await Assert.That(exception.Message).Contains("{% endfor %}");
        await Assert.That(exception.Message).Contains("same container");
    }

    [Test]
    public async Task IfInCellAndEndIfInBodyIsRejected()
    {
        using var template = BuildTemplate(
            BuildParagraph("Before"),
            new Table(
                new TableRow(
                    new TableCell(
                        BuildParagraph("{% if Flag %}"),
                        BuildParagraph("In the cell")))),
            BuildParagraph("Between"),
            BuildParagraph("{% endif %}"),
            BuildParagraph("After"));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<ConditionalTests.FlagModel>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("{% if Flag %}");
        await Assert.That(exception.Message).Contains("{% endif %}");
        await Assert.That(exception.Message).Contains("same container");
    }

    [Test]
    public async Task ElseInADifferentContainerIsRejected()
    {
        // Open and close are siblings; the branch tag between them is not. The chosen branch is
        // kept by walking siblings from its tag, so every tag of the block has to line up.
        using var template = BuildTemplate(
            BuildParagraph("{% if Flag %}"),
            BuildParagraph("Yes"),
            new Table(
                new TableRow(
                    new TableCell(
                        BuildParagraph("{% else %}")))),
            BuildParagraph("No"),
            BuildParagraph("{% endif %}"));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<ConditionalTests.FlagModel>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("{% if Flag %}");
        await Assert.That(exception.Message).Contains("{% else %}");
        await Assert.That(exception.Message).Contains("same container");
    }

    [Test]
    public async Task LoopInsideOneCellRenders()
    {
        using var template = BuildTemplate(
            new Table(
                new TableRow(
                    new TableCell(
                        BuildParagraph("{% for line in Lines %}"),
                        BuildParagraph("{{ line.Description }}"),
                        BuildParagraph("{% endfor %}")))));

        var texts = await RenderCellTexts(template, SampleData.Invoice());
        var expected = SampleData.Invoice().Lines.Select(_ => _.Description).ToList();
        await Assert.That(texts).IsEquivalentTo(expected);
    }

    // A range loop is a block like any other: where its tags sit is checked whatever it iterates.
    [Test]
    public async Task RangeForAndEndForInDifferentTableRowsIsRejected()
    {
        using var template = BuildTemplate(
            new Table(
                new TableRow(new TableCell(BuildParagraph("{% for i in (1..3) %}"))),
                new TableRow(new TableCell(BuildParagraph("{{ i }}"))),
                new TableRow(new TableCell(BuildParagraph("{% endfor %}")))));

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Invoice>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("{% for i in (1..3) %}");
        await Assert.That(exception.Message).Contains("same container");
    }

    // The bound is a liquid member of the collection, validated like any other reference.
    [Test]
    public async Task RangeLoopInsideOneCellRenders()
    {
        using var template = BuildTemplate(
            new Table(
                new TableRow(
                    new TableCell(
                        BuildParagraph("{% for i in (1..Lines.size) %}"),
                        BuildParagraph("Row {{ i }}"),
                        BuildParagraph("{% endfor %}")))));

        var texts = await RenderCellTexts(template, SampleData.Invoice());
        var expected = Enumerable.Range(1, SampleData.Invoice().Lines.Count).Select(_ => $"Row {_}").ToList();
        await Assert.That(texts).IsEquivalentTo(expected);
    }

    [Test]
    public async Task LoopRepeatingAWholeTableRenders()
    {
        using var template = BuildTemplate(
            BuildParagraph("{% for line in Lines %}"),
            new Table(
                new TableRow(
                    new TableCell(
                        BuildParagraph("{{ line.Description }}")))),
            BuildParagraph("{% endfor %}"));

        var invoice = SampleData.Invoice();
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);
        using var stream = new MemoryStream();
        await store.Render(invoice, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var tables = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Select(_ => _.InnerText).ToList();
        await Assert.That(tables).IsEquivalentTo(invoice.Lines.Select(_ => _.Description).ToList());
    }

    [Test]
    public async Task BlockInACellNestedInABodyLoopRenders()
    {
        // The rule is per block: the inner block's tags share the cell, the outer block's tags
        // share the body. Nothing requires the two blocks to share a container with each other.
        using var template = BuildTemplate(
            BuildParagraph("{% for line in Lines %}"),
            new Table(
                new TableRow(
                    new TableCell(
                        BuildParagraph("{% if line.Description %}"),
                        BuildParagraph("{{ line.Description }}"),
                        BuildParagraph("{% endif %}")))),
            BuildParagraph("{% endfor %}"));

        var invoice = SampleData.Invoice();
        var texts = await RenderCellTexts(template, invoice);
        await Assert.That(texts).IsEquivalentTo(invoice.Lines.Select(_ => _.Description).ToList());
    }

    static async Task<List<string>> RenderCellTexts(Stream template, Invoice model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!
            .Descendants<TableCell>()
            .SelectMany(_ => _.Elements<Paragraph>())
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();
    }

    static MemoryStream BuildTemplate(params OpenXmlElement[] blocks)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body(blocks);
            body.Append(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));
            mainPart.Document = new(body);
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph BuildParagraph(string text) =>
        new(
            new Run(
                new Text(text)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
}
