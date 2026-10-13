

// ReSharper disable PartialTypeWithSinglePart
public partial class UnusableLinkTests
{
    [AttributeUsage(AttributeTargets.Property)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class Model
    {
        public string Url { get; set; } = "";

        [Markdown]
        public string Md { get; set; } = "";
    }

    // Each of these used to reach new Uri and throw UriFormatException out of the render.
    [Test]
    [Arguments("http://")]
    [Arguments("https://:80")]
    [Arguments("http://host:99999999/x")]
    [Arguments("http://[bad")]
    [Arguments("http:///x")]
    [Arguments("x:y:z")]
    [Arguments("http://exa%mple.com/%zz")]
    [Arguments("ht!tp://a")]
    [Arguments("not a url")]
    public async Task BoundDestinationThatIsNotAUrlStaysText(string url)
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("[text]({{ Url }}) tail");

        using var document = await Render(
            store,
            new()
            {
                Url = url
            });

        var mainPart = document.MainDocumentPart!;
        await Assert.That(mainPart.Document!.Descendants<Hyperlink>().Any()).IsFalse();
        await Assert.That(mainPart.HyperlinkRelationships.Any()).IsFalse();
        await Assert.That(mainPart.Document.InnerText).Contains("text");
        await Assert.That(mainPart.Document.InnerText).EndsWith(" tail");
        await Assert.That(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document).Any()).IsFalse();
    }

    // The text followed by the destination in parentheses, which is what OpenXmlHtml writes for an
    // <a href> it cannot resolve.
    [Test]
    public async Task TextIsFollowedByTheDestination()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("See [the *site*](http://) now").Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("See the site (http://) now");
        await Assert.That(paragraph.Descendants<Hyperlink>().Any()).IsFalse();
        await Assert.That(paragraph.Descendants<RunStyle>().Any()).IsFalse();
    }

    [Test]
    public async Task TextThatIsTheDestinationIsNotRepeated()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("[http://](http://)").Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("http://");
    }

    [Test]
    public async Task EmptyTextLeavesTheDestination()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("See [](http://) now").Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("See http:// now");
    }

    // A relative destination has nothing to be relative to in a generated document, and whether a
    // rooted one parses as a file path depends on the platform, so none of them is linked.
    [Test]
    [Arguments("/docs/page")]
    [Arguments("//host/page")]
    [Arguments("page.html")]
    [Arguments("../page.html")]
    [Arguments("?query")]
    public async Task RelativeDestinationStaysText(string url)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstInline<LinkInline>($"[page]({url})"));

        await Assert.That(renderer.Top.CurrentRuns.OfType<Hyperlink>().Any()).IsFalse();
        await Assert.That(string.Concat(renderer.Top.CurrentRuns.Select(_ => _.InnerText))).IsEqualTo($"page ({url})");
        await Assert.That(renderer.MainPart.HyperlinkRelationships.Any()).IsFalse();
    }

    [Test]
    [Arguments("https://example.com/a?b=c#d")]
    [Arguments("mailto:someone@example.com")]
    [Arguments("tel:+61400000000")]
    [Arguments("file:///C:/path/file.txt")]
    public async Task AbsoluteDestinationIsStillLinked(string url)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstInline<LinkInline>($"[page]({url})"));

        var link = (Hyperlink)renderer.Top.CurrentRuns.Single();
        var relationship = renderer.MainPart.HyperlinkRelationships.Single();
        await Assert.That(link.Id?.Value).IsEqualTo("rIdLink1");
        await Assert.That(relationship.Id).IsEqualTo("rIdLink1");
        await Assert.That(relationship.Uri.OriginalString).IsEqualTo(url);
    }

    // An unusable destination takes no relationship id, so the links after it number as if it were
    // not there.
    [Test]
    public async Task UnusableDestinationTakesNoRelationshipId()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("[a]({{ Url }}) [b](https://example.com) [c](page.html) [d](https://example.org)");

        var first = await RenderBytes(store);
        var second = await RenderBytes(store);
        await Assert.That(first).IsEquivalentTo(second);

        using var document = WordprocessingDocument.Open(new MemoryStream(first), false);
        await Assert.That(document.MainDocumentPart!.HyperlinkRelationships.Select(_ => _.Id).Order())
            .IsEquivalentTo(["rIdLink1", "rIdLink2"]);
    }

    // The bare form is a LinkInline built by BoundAutoLinkParser and the angle form an
    // AutolinkInline, so the two take different renderers.
    [Test]
    [Arguments("See http://{{ Url }} now")]
    [Arguments("See <http://{{ Url }}> now")]
    public async Task BoundAutolinkThatIsNotAUrlStaysText(string template)
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>(template);

        using var document = await Render(
            store,
            new()
            {
                Url = "[bad"
            });

        var mainPart = document.MainDocumentPart!;
        await Assert.That(mainPart.Document!.Descendants<Hyperlink>().Any()).IsFalse();
        await Assert.That(mainPart.Document.InnerText).IsEqualTo("See http://[bad now");
    }

    // An image source goes through OpenXmlHtml rather than a hyperlink relationship.
    [Test]
    public async Task ImageWithUnusableSourceDoesNotThrow()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("![alt]({{ Url }}) tail");

        using var document = await Render(
            store,
            new()
            {
                Url = "http://"
            });

        await Assert.That(document.MainDocumentPart!.Document!.InnerText).EndsWith(" tail");
    }

    [Test]
    public async Task MarkdownMemberInDocxTemplate()
    {
        using var template = DocxTemplateBuilder.Build("Intro\n\n{{ Md }}\n\nOutro");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var document = await Render(
            store,
            new()
            {
                Md = "[text](http://) and [site](https://example.com) tail"
            });

        var mainPart = document.MainDocumentPart!;
        await Assert.That(mainPart.Document!.InnerText).IsEqualTo("Introtext (http://) and site tailOutro");
        await Assert.That(mainPart.Document.Descendants<Hyperlink>().Single().InnerText).IsEqualTo("site");
        await Assert.That(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document).Any()).IsFalse();
    }

    static async Task<byte[]> RenderBytes(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(
            new Model
            {
                Url = "http://"
            },
            stream);
        return stream.ToArray();
    }

    static async Task<WordprocessingDocument> Render(TemplateStore store, Model model)
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }
}
