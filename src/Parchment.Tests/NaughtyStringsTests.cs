// ReSharper disable PartialTypeWithSinglePart
public partial class NaughtyStringsTests
{
    [ParchmentBindable]
    public partial class NaughtyModel
    {
        public required string Single { get; init; }
        public required IReadOnlyList<NaughtyItem> Items { get; init; }
    }

    public class NaughtyItem
    {
        public required string Value { get; init; }
    }

    static NaughtyModel BuildModel() =>
        new()
        {
            Single = string.Join(" | ", TheNaughtyStrings.All),
            Items = TheNaughtyStrings.All
                .Select(_ => new NaughtyItem {Value = _})
                .ToList()
        };

    // Substitution into runs, so these strings were never anything but text here.
    [Test]
    public async Task Docx()
    {
        using var stream = await RenderDocx("");
        await Verify(stream, "docx");
    }

    // The docx flow is where the two markup filters differ: each produces its own token type, and
    // the value is parsed by a different parser and spliced in structurally.
    [Test]
    public async Task DocxFilterMarkdown()
    {
        using var stream = await RenderDocx(" | markdown");
        await Verify(stream, "docx");
    }

    [Test]
    public async Task DocxFilterHtml()
    {
        using var stream = await RenderDocx(" | html");
        await Verify(stream, "docx");
    }

    static async Task<MemoryStream> RenderDocx(string filter)
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Single: {{ Single_FILTER_ }}

            {% for item in Items %}

            {{ item.Value_FILTER_ }}

            {% endfor %}
            """.Replace("_FILTER_", filter));

        var store = new TemplateStore();
        store.RegisterDocxTemplate<NaughtyModel>(template);

        var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;
        return stream;
    }

    // What a bound value is by default: text. Every string arrives in the document as itself,
    // whatever it happens to look like to a markdown parser.
    [Test]
    public async Task Markdown()
    {
        using var stream = await RenderMarkdown("");
        await Verify(stream, "docx");
    }

    // The hostile path: every string handed to Markdig as source, so unbalanced emphasis, stray
    // pipes, generic attributes and the corpus's xss payloads are all syntax. Nothing here asserts
    // the output is sensible — for most of these inputs there is no sensible output — only that the
    // pipeline survives them and lands somewhere stable, which is what the snapshot pins.
    //
    // One snapshot covers all three opt-outs because in this flow they are one operation; see
    // OptOutFiltersAreEquivalent.
    [Test]
    public async Task MarkdownUnescaped()
    {
        using var stream = await RenderMarkdown(" | raw");
        await Verify(stream, "docx");
    }

    // Every string handed to the html converter instead, which is a different parser reading the
    // same bytes — the corpus's unclosed tags, bogus attributes and xss payloads are markup here.
    [Test]
    public async Task MarkdownFilterHtml()
    {
        using var stream = await RenderMarkdown(" | html");
        await Verify(stream, "docx");
    }

    // `raw` and `markdown` are one operation in a markdown template: both mean "do not encode", and
    // markdown source written into markdown is already what it will be parsed as. So one snapshot
    // covers both, with this holding them to it rather than pinning 32 pages of the same document
    // twice.
    //
    // `html` is not in that group, and the second assertion is what keeps it out. Written straight
    // into the source it would be classified by Markdig rather than converted as html — which is
    // markdown's answer, not html's — so it takes a marker and is converted after the parse. If
    // that ever regresses to a pass-through, this is what says so.
    [Test]
    public async Task PassThroughFiltersAreEquivalent()
    {
        var raw = await RenderMarkdownBody(" | raw");

        await Assert.That(await RenderMarkdownBody(" | markdown")).IsEqualTo(raw);
        await Assert.That(await RenderMarkdownBody(" | html")).IsNotEqualTo(raw);
    }

    // The body rather than the stream, because the packaging around it carries part ordering that
    // is not what the filters are being compared on. Relationship ids go the same way: the OpenXml
    // SDK mints a fresh one per hyperlink per render — the corpus contains urls, which autolink —
    // so two identical documents never share them. Verify scrubs them for the same reason, which is
    // why the snapshots of these matched while their bytes did not.
    static async Task<string> RenderMarkdownBody(string filter)
    {
        using var stream = await RenderMarkdown(filter);
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart!.Document!.Body!.OuterXml;
        return System.Text.RegularExpressions.Regex.Replace(body, "r:id=\"[^\"]*\"", "r:id=\"_\"");
    }

    static async Task<MemoryStream> RenderMarkdown(string filter)
    {
        // Spliced rather than interpolated: liquid's own {{ }} collides with every interpolation
        // form a raw string offers.
        var markdownSource =
            """
            # {{ Single_FILTER_ }}

            {% for item in Items %}
            - {{ item.Value_FILTER_ }}
            {% endfor %}
            """.Replace("_FILTER_", filter);

        using var styleSource = DocxTemplateBuilder.Build();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<NaughtyModel>(
            markdownSource,
            styleSource: styleSource);

        var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;
        return stream;
    }
}
