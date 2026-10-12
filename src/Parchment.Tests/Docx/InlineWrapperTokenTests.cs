// An inline wrapper sits between a paragraph and its runs. A token typed inside one, or split
// between one and the runs beside it, is still the paragraph's token. w:smartTag is the odd one
// out: the Open XML SDK has no class for it, so it and everything inside it load as unknown
// elements.
// ReSharper disable PartialTypeWithSinglePart
public partial class InlineWrapperTokenTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [ParchmentBindable]
    public partial class Model
    {
        public required string Title { get; init; }

        public bool Show { get; init; }

        public List<string> Tags { get; init; } = [];

        [Html]
        public string Body { get; init; } = "";
    }

    const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string smartTagOpen = """<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">""";
    const string smartTagClose = "</w:smartTag>";

    [Test]
    [Arguments("smartTag", smartTagOpen, smartTagClose)]
    [Arguments("smartTag", smartTagOpen + """<w:smartTagPr><w:attr w:name="ProductID" w:val="x"/></w:smartTagPr>""", smartTagClose)]
    [Arguments("smartTag", smartTagOpen + smartTagOpen, smartTagClose + smartTagClose)]
    [Arguments("smartTag", """<w:hyperlink w:anchor="x">""" + smartTagOpen, smartTagClose + "</w:hyperlink>")]
    [Arguments("hyperlink", smartTagOpen + """<w:hyperlink w:anchor="x">""", "</w:hyperlink>" + smartTagClose)]
    [Arguments("hyperlink", """<w:hyperlink w:anchor="x">""", "</w:hyperlink>")]
    [Arguments("customXml", """<w:customXml w:uri="u" w:element="e">""", "</w:customXml>")]
    [Arguments("dir", """<w:dir w:val="ltr">""", "</w:dir>")]
    [Arguments("bdo", """<w:bdo w:val="ltr">""", "</w:bdo>")]
    [Arguments("fldSimple", """<w:fldSimple w:instr=" QUOTE ">""", "</w:fldSimple>")]
    [Arguments("ins", """<w:ins w:id="3" w:author="a" w:date="2026-01-01T00:00:00Z">""", "</w:ins>")]
    [Arguments("moveTo", """<w:moveTo w:id="3" w:author="a" w:date="2026-01-01T00:00:00Z">""", "</w:moveTo>")]
    [Arguments("sdt", """<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>""", "</w:sdtContent></w:sdt>")]
    public async Task TokenInsideAWrapperIsSubstituted(string wrapper, string open, string close)
    {
        using var template = Build(P(open + R("x {{ Title }} y") + close));

        using var result = await Render(
            template,
            new()
            {
                Title = "Al"
            });

        var paragraph = Paragraphs(result).Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("x Al y");
        await Assert.That(paragraph.Descendants().Any(_ => _.LocalName == wrapper)).IsTrue();
    }

    [Test]
    public async Task TokenSplitBetweenASmartTagAndItsNeighboursIsSubstituted()
    {
        using var template = Build(P(R("Dear {{ Ti") + smartTagOpen + R("tl") + smartTagClose + R("e }}, welcome")));

        using var result = await Render(
            template,
            new()
            {
                Title = "Al"
            });

        await Assert.That(Paragraphs(result).Single().InnerText).IsEqualTo("Dear Al, welcome");
    }

    [Test]
    public async Task TokenSplitBetweenNestedSmartTagsIsSubstituted()
    {
        using var template = Build(P(smartTagOpen + R("{{ ") + smartTagOpen + R("Title") + smartTagClose + R(" }}") + smartTagClose + R(" and {{ Title }}")));

        using var result = await Render(
            template,
            new()
            {
                Title = "Alexandria"
            });

        await Assert.That(Paragraphs(result).Single().InnerText).IsEqualTo("Alexandria and Alexandria");
    }

    // The smart tag stays in the output, holding the substituted run with its formatting.
    [Test]
    public async Task SmartTagAndRunFormattingSurviveTheSubstitution()
    {
        using var template = Build(P(smartTagOpen + "<w:r><w:rPr><w:b/></w:rPr><w:t>{{ Title }}</w:t></w:r>" + smartTagClose));

        using var result = await Render(
            template,
            new()
            {
                Title = "Al"
            });

        var tag = Paragraphs(result).Single().ChildElements.Single(_ => _.LocalName == "smartTag");
        await Assert.That(tag.OuterXml).Contains("urn:schemas-microsoft-com:office:smarttags");
        await Assert.That(tag.OuterXml).Contains("<w:b");
        await Assert.That(tag.InnerText).IsEqualTo("Al");
    }

    [Test]
    public async Task MultiLineValueInsideASmartTagBecomesALineBreak()
    {
        using var template = Build(P(smartTagOpen + R("{{ Title }}") + smartTagClose));

        using var result = await Render(
            template,
            new()
            {
                Title = "one\ntwo"
            });

        var paragraph = Paragraphs(result).Single();
        await Assert.That(paragraph.OuterXml).Contains("<w:br");
        await Assert.That(paragraph.InnerText).IsEqualTo("onetwo");
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task BlockTagsInsideSmartTagsAreSeen(bool show)
    {
        using var template = Build(
            P(smartTagOpen + R("{% if Show %}") + smartTagClose),
            P(R("shown to ") + smartTagOpen + R("{{ Title }}") + smartTagClose),
            P(R("{% end") + smartTagOpen + R("if") + smartTagClose + R(" %}")),
            P(R("after")));

        using var result = await Render(
            template,
            new()
            {
                Title = "Al",
                Show = show
            });

        var texts = Paragraphs(result).Select(_ => _.InnerText).ToList();
        if (show)
        {
            await Assert.That(texts).IsEquivalentTo(["shown to Al", "after"]);
            return;
        }

        await Assert.That(texts).IsEquivalentTo(["after"]);
    }

    [Test]
    public async Task LoopBodyInsideASmartTagIsRepeated()
    {
        using var template = Build(
            P(smartTagOpen + R("{% for tag in Tags %}") + smartTagClose),
            P(R("- ") + smartTagOpen + R("{{ tag }}") + smartTagClose),
            P(R("{% endfor %}")));

        using var result = await Render(
            template,
            new()
            {
                Title = "Al",
                Tags = ["one", "two"]
            });

        await Assert.That(Paragraphs(result).Select(_ => _.InnerText).ToList()).IsEquivalentTo(["- one", "- two"]);
    }

    // A block-level value splits its host into a head and a tail around the produced blocks.
    [Test]
    public async Task SplicedTokenInsideASmartTagKeepsTheTextAroundIt()
    {
        using var template = Build(P(R("Intro ") + smartTagOpen + R("mid {{ Body }} dle") + smartTagClose + R(" outro")));

        using var result = await Render(
            template,
            new()
            {
                Title = "Al",
                Body = "<p>First</p><p>Second</p>"
            });

        var text = string.Concat(Paragraphs(result).Select(_ => _.InnerText + "|"));
        await Assert.That(text).IsEqualTo("Intro mid |First|Second| dle outro|");
    }

    [Test]
    public async Task UnknownMemberInsideASmartTagFailsRegistration()
    {
        using var template = Build(P(smartTagOpen + R("{{ Missing }}") + smartTagClose));
        var store = new TemplateStore();

        await Assert.That(() => store.RegisterDocxTemplate<Model>(template)).Throws<ParchmentRegistrationException>();
    }

    [Test]
    public async Task TokenInsideASmartTagInAHeaderIsSubstituted()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            header.Header = new(new Paragraph(P(smartTagOpen + R("head {{ Title }}") + smartTagClose)));
            main.Document = new(
                new Body(
                    new Paragraph(),
                    new SectionProperties(
                        new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdHeader"
                        })));
        }

        stream.Position = 0;
        using var result = await Render(
            stream,
            new()
            {
                Title = "Al"
            });

        var rendered = result.MainDocumentPart!.HeaderParts.Single().Header!;
        await Assert.That(rendered.InnerText).IsEqualTo("head Al");
    }

    static string P(string content) =>
        $"""<w:p xmlns:w="{wordNamespace}">{content}</w:p>""";

    static string R(string text) =>
        $"""<w:r><w:t xml:space="preserve">{text}</w:t></w:r>""";

    static List<Paragraph> Paragraphs(WordprocessingDocument document) =>
        document.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .ToList();

    static async Task<WordprocessingDocument> Render(Stream template, Model model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var output = new MemoryStream();
        await store.Render(model, output);
        output.Position = 0;
        return WordprocessingDocument.Open(output, false);
    }

    static MemoryStream Build(params string[] paragraphs)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(paragraphs.Select(_ => new Paragraph(_))));
        }

        stream.Position = 0;
        return stream;
    }
}
