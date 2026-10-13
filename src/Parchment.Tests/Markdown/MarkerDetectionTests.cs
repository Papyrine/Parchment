// ReSharper disable PartialTypeWithSinglePart
// What counts as a [TOC] or [SECTION] marker, and where one can come from. The markdown flow turns
// most paragraphs away before reading their text and skips its section pass when the walk wrote no
// marker, so each case here is one of the routes a marker can take that those shortcuts must not
// lose - and each is pinned as it behaved before they existed.
public partial class MarkerDetectionTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public string Text = "";
        public MarkdownToken Token = new("");
        public OpenXmlToken Raw = OpenXmlToken.Empty;
        public bool Show;
        public List<string> Items = [];
    }

    const string landscape = "[SECTION]{orientation=landscape}";

    // A bound value is data. One that reads as the marker prints as that text, the way any other
    // markdown in a value does - MarkdownEncoder writes its opening bracket as a character
    // reference, and a marker is a paragraph of nothing but literals.
    [Test]
    public async Task BoundValueReadingAsTocStaysText()
    {
        var result = await Inspect(
            Between("{{ Text }}"),
            new()
            {
                Text = "[TOC]"
            });

        await Assert.That(result.Toc).IsEqualTo(0);
        await Assert.That(result.Sections).IsEqualTo(1);
        await Assert.That(result.Text).Contains("[TOC]");
    }

    // Escaping takes the braces of the attribute with it as well.
    [Test]
    public async Task BoundValueReadingAsSectionStaysText()
    {
        var result = await Inspect(
            Between("{{ Text }}"),
            new()
            {
                Text = landscape
            });

        await Assert.That(result.Sections).IsEqualTo(1);
        await Assert.That(result.Text).Contains(landscape);
    }

    // Nor does the template supplying the attribute make it one: the attribute lands on the
    // paragraph, and the paragraph is still a value.
    [Test]
    public async Task BoundValueWithTheTemplatesAttributeStaysText()
    {
        var result = await Inspect(
            Between("{{ Text }}{orientation=landscape}"),
            new()
            {
                Text = "[SECTION]"
            });

        await Assert.That(result.Sections).IsEqualTo(1);
        await Assert.That(result.Text).Contains("[SECTION]");
    }

    [Test]
    public async Task MarkdownFilterValueCanBeEitherMarker()
    {
        var toc = await Inspect(
            Between("{{ Text | markdown }}"),
            new()
            {
                Text = "[TOC]{levels=2}"
            });
        var section = await Inspect(
            Between("{{ Text | markdown }}"),
            new()
            {
                Text = landscape
            });

        await Assert.That(toc.Toc).IsEqualTo(1);
        await Assert.That(toc.Text).Contains("\"1-2\"");
        await Assert.That(section.Sections).IsEqualTo(2);
    }

    [Test]
    public async Task MarkdownTokenValueCanBeEitherMarker()
    {
        var toc = await Inspect(
            Between("{{ Token }}"),
            new()
            {
                Token = new("[TOC]")
            });
        var section = await Inspect(
            Between("{{ Token }}"),
            new()
            {
                Token = new($"Before\n\n{landscape}\n\nAfter")
            });

        await Assert.That(toc.Toc).IsEqualTo(1);
        await Assert.That(section.Sections).IsEqualTo(2);
        await Assert.That(section.Text).IsEqualTo("Intro|Before||After|End");
    }

    // The template source holds one marker and the rendered document two, so nothing read off the
    // source at registration could say how many there are.
    [Test]
    public async Task MarkerInsideALoopIsASectionPerIteration()
    {
        var result = await Inspect(
            $$$"""
               {% for item in Items %}

               {{ item }}

               {{{landscape}}}

               {% endfor %}

               End
               """,
            new()
            {
                Items = ["a", "b"]
            });

        await Assert.That(result.Sections).IsEqualTo(3);
        await Assert.That(result.Text).IsEqualTo("a||b||End");
    }

    [Test]
    [Arguments(true, 2, 1)]
    [Arguments(false, 1, 0)]
    public async Task MarkerInsideAConditionalFollowsTheCondition(bool show, int sections, int toc)
    {
        var result = await Inspect(
            $$$"""
               Intro

               {% if Show %}

               {{{landscape}}}

               [TOC]

               {% endif %}

               End
               """,
            new()
            {
                Show = show
            });

        await Assert.That(result.Sections).IsEqualTo(sections);
        await Assert.That(result.Toc).IsEqualTo(toc);
    }

    [Test]
    [Arguments("   [SECTION]{orientation=landscape}", "  [TOC]")]
    [Arguments(@"\[SECTION\]{orientation=landscape}", @"\[TOC\]")]
    public async Task IndentedOrEscapedMarkerIsStillAMarker(string section, string toc)
    {
        var result = await Inspect($"Intro\n\n{section}\n\n{toc}\n\nEnd", new());

        await Assert.That(result.Sections).IsEqualTo(2);
        await Assert.That(result.Toc).IsEqualTo(1);
    }

    [Test]
    [Arguments("[TOC] and more")]
    [Arguments("See [TOC]")]
    [Arguments("*[TOC]*")]
    [Arguments("[toc]")]
    [Arguments("- [TOC]")]
    [Arguments("[SECTION] x {orientation=landscape}")]
    [Arguments("See [SECTION]{orientation=landscape}")]
    [Arguments("**[SECTION]**{orientation=landscape}")]
    [Arguments("[section]{orientation=landscape}")]
    [Arguments("- [SECTION]{orientation=landscape}")]
    [Arguments("> [SECTION]{orientation=landscape}")]
    public async Task ParagraphThatOnlyResemblesAMarkerStaysText(string paragraph)
    {
        var result = await Inspect(Between(paragraph), new());

        await Assert.That(result.Sections).IsEqualTo(1);
        await Assert.That(result.Toc).IsEqualTo(0);
        await Assert.That(result.Text).Contains("[");
    }

    [Test]
    public async Task SectionMarkerBesideAnHtmlBlockIsASection()
    {
        var result = await Inspect($"<div>html</div>\n\n{landscape}\n\nEnd", new());

        await Assert.That(result.Sections).IsEqualTo(2);
    }

    // Not a supported way to ask for a section, but a paragraph carrying a w:sectPr of the shape a
    // marker writes is read as one wherever it came from, and a token is caller code the markdown
    // walk never saw. So the section pass still runs for a render that parked one.
    [Test]
    public async Task ParagraphFromAnOpenXmlTokenCarryingASectionIsRewritten()
    {
        using var stream = await Render(
            Between("{{ Raw }}"),
            new()
            {
                Raw = new(_ =>
                [
                    new Paragraph(
                        new ParagraphProperties(
                            new SectionProperties(
                                new PageSize
                                {
                                    Orient = PageOrientationValues.Landscape
                                })))
                ])
            });
        using var doc = WordprocessingDocument.Open(stream, false);
        var sections = doc.MainDocumentPart!.Document!.Body!
            .Descendants<SectionProperties>()
            .ToList();

        await Assert.That(sections.Count).IsEqualTo(2);
        // Rewritten from the style source rather than left as the bare w:pgSz the token wrote.
        await Assert.That(sections[0].GetFirstChild<PageMargin>()).IsNotNull();
        await Assert.That(sections[1].GetFirstChild<PageSize>()!.Orient!.Value).IsEqualTo(PageOrientationValues.Landscape);
    }

    // The same document with and without a marker differs by the marker's paragraph and nothing
    // else, which is what a skipped pass having been a no-op looks like.
    [Test]
    public async Task MarkerFreeRenderMatchesItsMarkedTwinApartFromTheMarker()
    {
        var plain = await Inspect("One\n\nTwo {#two}\n\n| A | B |\n| --- | --- |\n| c | d |", new());
        var marked = await Inspect($"One\n\n{landscape}\n\nTwo {{#two}}\n\n| A | B |\n| --- | --- |\n| c | d |", new());

        await Assert.That(plain.Sections).IsEqualTo(1);
        await Assert.That(marked.Sections).IsEqualTo(2);
        await Assert.That(marked.Text).IsEqualTo(plain.Text.Replace("One|", "One||"));
    }

    static string Between(string paragraph) =>
        $"Intro\n\n{paragraph}\n\nEnd";

    static async Task<(int Sections, int Toc, string Text)> Inspect(string markdown, Model model)
    {
        using var stream = await Render(markdown, model);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        return (
            body.Descendants<SectionProperties>().Count(),
            body.Descendants<FieldCode>().Count(_ => _.Text.Contains("TOC")),
            string.Join('|', body.Elements<Paragraph>().Select(_ => _.InnerText)));
    }

    static async Task<MemoryStream> Render(string markdown, Model model)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }
}
