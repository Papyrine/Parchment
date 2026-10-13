// ReSharper disable PartialTypeWithSinglePart

// The same model value has to print as the same text whichever flow renders it. The docx flow
// reads a token's value with ToStringValue and the markdown flow lets Fluid write it, and Fluid's
// two paths disagree for a number; a TimeSpan and a TimeOnly were left to Fluid's defaults in both.
public partial class ValueFormattingTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public decimal Amount { get; init; } = 1234.50m;
        public decimal Whole { get; init; } = 10.00m;
        public decimal Small { get; init; } = 12.5m;
        public int Count { get; init; } = 3;
        public double Ratio { get; init; } = 2.7;
        public TimeSpan Span { get; init; } = new(1, 2, 3, 4);
        public TimeSpan Short { get; init; } = new(0, 0, 90);
        public TimeSpan? NoSpan { get; init; }
        public Time Time { get; init; } = new(13, 5, 6);
        public Time? MaybeTime { get; init; } = new(7, 8, 9);
        public IReadOnlyList<Line> Lines { get; init; } = [new(), new()];
    }

    [ParchmentBindable]
    public partial class Line
    {
        public decimal Price { get; init; } = 2.50m;
    }

    [Test]
    [Arguments("{{ Amount }}", "1234.50")]
    [Arguments("{{ Whole }}", "10.00")]
    [Arguments("{{ Small }}", "12.5")]
    [Arguments("{{ Count }}", "3")]
    [Arguments("{{ Ratio }}", "2.7")]
    [Arguments("{{ Amount | plus: 1 }}", "1235.50")]
    [Arguments("{{ Span }}", "1.02:03:04")]
    [Arguments("{{ Short }}", "00:01:30")]
    [Arguments("[{{ NoSpan }}]", "[]")]
    [Arguments("{{ Time }}", "13:05:06")]
    [Arguments("{{ MaybeTime }}", "07:08:09")]
    [Arguments("{{ Time | date: '%H-%M-%S' }}", "13-05-06")]
    public async Task BothFlowsPrintTheSameText(string template, string expected)
    {
        using var docx = DocxTemplateBuilder.Build(template);
        var docxStore = new TemplateStore();
        docxStore.RegisterDocxTemplate<Doc>(docx);

        using var styleSource = DocxTemplateBuilder.Build();
        var markdownStore = new TemplateStore();
        markdownStore.RegisterMarkdownTemplate<Doc>(template, styleSource);

        await Assert.That(await RenderText(docxStore)).IsEqualTo(expected);
        await Assert.That(await RenderText(markdownStore)).IsEqualTo(expected);
    }

    // A markdown template is one liquid document, so a number can be written from inside a block
    // or from a variable, and has to stay a number until then for the comparison to work.
    [Test]
    [Arguments("{% if Amount > 1000 %}{{ Amount }}{% endif %}", "1234.50")]
    [Arguments("{% for line in Lines %}{{ line.Price }};{% endfor %}", "2.50;2.50;")]
    [Arguments("{% assign copy = Whole %}{{ copy }}", "10.00")]
    [Arguments("{% capture text %}{{ Whole }}{% endcapture %}{{ text }}", "10.00")]
    public async Task MarkdownKeepsTheScaleInsideBlocks(string template, string expected)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(template, styleSource);

        await Assert.That(await RenderText(store)).IsEqualTo(expected);
    }

    static async Task<string> RenderText(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(new Doc(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
