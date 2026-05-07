public class FilterTests
{
    public class Doc
    {
        public required string Body { get; init; }
        public required IReadOnlyList<string> Items { get; init; }
    }

    [Test]
    public async Task EscapeXml_EscapesAllSpecialChars()
    {
        using var template = DocxTemplateBuilder.Build("{{ Body | escape_xml }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("escape-xml", template);

        using var stream = new MemoryStream();
        await store.Render(
            "escape-xml",
            new Doc
            {
                Body = "<a>\"hi\" & 'bye' > done</a>",
                Items = []
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var text = string.Concat(doc.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(_ => _.Text));
        await Assert.That(text).IsEqualTo("&lt;a&gt;&quot;hi&quot; &amp; &apos;bye&apos; &gt; done&lt;/a&gt;");
    }

    [Test]
    public async Task EscapeXml_PassesThroughBenignText()
    {
        using var template = DocxTemplateBuilder.Build("{{ Body | escape_xml }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("escape-xml-benign", template);

        using var stream = new MemoryStream();
        await store.Render(
            "escape-xml-benign",
            new Doc
            {
                Body = "no special chars here",
                Items = []
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var text = string.Concat(doc.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(_ => _.Text));
        await Assert.That(text).IsEqualTo("no special chars here");
    }

    [Test]
    public async Task EscapeXmlChainedWithUpcase_BothApplied()
    {
        // Filter chains compose left-to-right: escape_xml then upcase. Both filters return
        // FluidValue (escape_xml: StringValue, upcase: StringValue), so the chain works.
        using var template = DocxTemplateBuilder.Build("{{ Body | escape_xml | upcase }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("escape-upcase", template);

        using var stream = new MemoryStream();
        await store.Render(
            "escape-upcase",
            new Doc
            {
                Body = "<x>hi & bye</x>",
                Items = []
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var text = string.Concat(doc.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(_ => _.Text));
        await Assert.That(text).IsEqualTo("&LT;X&GT;HI &AMP; BYE&LT;/X&GT;");
    }

    [Test]
    public async Task BulletListFilter_EmptyEnumerable_ProducesNoListParagraphs()
    {
        // Document the bullet_list filter behavior on an empty IEnumerable<string>: no list
        // paragraphs are produced. The host paragraph (which contained the solo token) is
        // structurally replaced by zero elements.
        using var template = DocxTemplateBuilder.Build(
            """
            Header

            {{ Items | bullet_list }}

            Footer
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("bullet-empty", template);

        using var stream = new MemoryStream();
        await store.Render(
            "bullet-empty",
            new Doc
            {
                Body = "",
                Items = []
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        var listParagraphs = paragraphs.Where(_ =>
                _.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph")
            .ToList();
        await Assert.That(listParagraphs).IsEmpty();
    }

    [Test]
    public async Task BulletListFilter_DefaultsToListParagraph_WhenHostHasNoPStyle()
    {
        // Baseline: a host paragraph without an explicit pStyle gets the historical
        // "ListParagraph" default for produced bullet paragraphs.
        using var template = DocxTemplateBuilder.Build("{{ Items | bullet_list }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("bullet-default", template);

        using var stream = new MemoryStream();
        await store.Render(
            "bullet-default",
            new Doc
            {
                Body = "",
                Items = ["a", "b"]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var styles = BulletParagraphStyles(doc);
        await Assert.That(styles).IsEquivalentTo(["ListParagraph", "ListParagraph"]);
    }

    [Test]
    public async Task BulletListFilter_InheritsHostPStyle()
    {
        // When the host paragraph carries a pStyle (e.g. a table cell using "TBLText"), the
        // produced bullet paragraphs should adopt that style so they pick up the surrounding
        // font instead of falling back to "ListParagraph" / Normal. The bullet glyph + indent
        // come from the abstractNum level, so dropping "ListParagraph" is safe.
        using var template = BuildTemplateWithStyledHost(
            "{{ Items | bullet_list }}",
            hostStyle: "TBLText");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("bullet-inherit", template);

        using var stream = new MemoryStream();
        await store.Render(
            "bullet-inherit",
            new Doc
            {
                Body = "",
                Items = ["a", "b"]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var styles = BulletParagraphStyles(doc);
        await Assert.That(styles).IsEquivalentTo(["TBLText", "TBLText"]);
    }

    [Test]
    public async Task NumberedListFilter_InheritsHostPStyle()
    {
        // Numbered lists go through the same code path as bullet lists, so the inheritance
        // behavior is identical. Cover it explicitly to lock the contract for both.
        using var template = BuildTemplateWithStyledHost(
            "{{ Items | numbered_list }}",
            hostStyle: "TBLText");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>("numbered-inherit", template);

        using var stream = new MemoryStream();
        await store.Render(
            "numbered-inherit",
            new Doc
            {
                Body = "",
                Items = ["a", "b"]
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var styles = BulletParagraphStyles(doc);
        await Assert.That(styles).IsEquivalentTo(["TBLText", "TBLText"]);
    }

    static List<string> BulletParagraphStyles(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Where(_ => _.ParagraphProperties?.NumberingProperties != null)
            .Select(_ => _.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "")
            .ToList();

    static MemoryStream BuildTemplateWithStyledHost(string tokenText, string hostStyle)
    {
        var template = DocxTemplateBuilder.Build(tokenText);
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
                .First(_ => _.InnerText.Contains("{{"));
            paragraph.ParagraphProperties ??= new();
            paragraph.ParagraphProperties.PrependChild(
                new ParagraphStyleId
                {
                    Val = hostStyle
                });
        }
        template.Position = 0;
        return template;
    }
}
