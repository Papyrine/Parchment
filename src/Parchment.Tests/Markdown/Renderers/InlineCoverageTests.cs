/// <summary>
/// Markdig writes nothing for a leaf inline no renderer claims, and says nothing about it either. So
/// text the parser had moved into a node of its own - an entity, an angle quote, a bracket that
/// never closed - was absent from the document with no error anywhere. These pin that every inline
/// the configured pipeline can produce is rendered, whatever it sits inside.
/// </summary>
public class InlineCoverageTests
{
    [Test]
    public async Task EveryInlineThePipelineCanProduceHasARenderer()
    {
        var renderer = RendererHarness.BuildRenderer();

        // The core inlines plus those of each extension the pipeline turns on - an extension keeps
        // its inlines in its own namespace. Turning another one on therefore fails here until what
        // it produces is rendered, and so does a Markdig that grows a new core inline.
        var namespaces = MarkdigPipeline.Pipeline.Extensions
            .Select(_ => _.GetType().Namespace)
            .Append(typeof(Inline).Namespace)
            .ToHashSet();

        var unclaimed = typeof(Inline).Assembly
            .GetTypes()
            .Where(_ => _.IsAssignableTo(typeof(Inline)) &&
                        !_.IsAbstract &&
                        namespaces.Contains(_.Namespace))
            .Where(type => !renderer.ObjectRenderers.Any(_ => _.Accept(renderer, type)))
            .Select(_ => _.Name)
            .ToList();

        // A bare ContainerInline is only ever the root of a block's inline content, which is walked
        // child by child rather than rendered as a node - and Markdig writes the children of a
        // container nothing claims, so that one loses nothing either way.
        await Assert.That(unclaimed).IsEquivalentTo([nameof(ContainerInline)]);
    }

    static readonly (string Markdown, string Text)[] snippets =
    [
        ("AT&amp;T", "AT&T"),
        ("&copy; &#169; &lt;tag&gt;", "\u00A9 \u00A9 <tag>"),
        ("<<quoted>>", "\u00ABquoted\u00BB"),
        ("see [note", "see [note"),
        ("an ![image", "an ![image"),
        ("P&amp;L [open <<x>> end", "P&L [open \u00ABx\u00BB end")
    ];

    [Test]
    [Arguments("{0}", "{0}")]
    [Arguments("before {0}", "before {0}")]
    [Arguments("**{0}**", "{0}")]
    [Arguments("_{0}_", "{0}")]
    [Arguments("~~{0}~~", "{0}")]
    [Arguments("==**{0}**==", "{0}")]
    [Arguments("<u>{0}</u>", "{0}")]
    [Arguments("# {0}", "{0}")]
    [Arguments("### {0} {{.Title}}", "{0}")]
    [Arguments("- {0}", "{0}")]
    [Arguments("1. {0}", "{0}")]
    [Arguments("- outer\n  - **{0}**", "outer{0}")]
    [Arguments("> {0}", "{0}")]
    [Arguments("| Head |\n| --- |\n| {0} |", "Head{0}")]
    [Arguments("| {0} |\n| --- |\n| body |", "{0}body")]
    [Arguments("| Head | Other |\n| --- | --- |\n| **{0}** | after |", "HeadOther{0}after")]
    public async Task RendersInsideEveryContainer(string container, string expected)
    {
        foreach (var (markdown, text) in snippets)
        {
            var elements = RendererHarness.RenderMarkdown(string.Format(container, markdown));

            await Assert.That(string.Concat(elements.Select(_ => _.InnerText)))
                .IsEqualTo(string.Format(expected, text));
        }
    }

    // A bracket cannot be left open inside link text - one that is makes the outer bracket the open
    // one instead - so the link carries the other two.
    [Test]
    [Arguments("[AT&amp;T](https://example.com)", "AT&T")]
    [Arguments("[<<quoted>> &copy;](https://example.com)", "\u00ABquoted\u00BB \u00A9")]
    [Arguments("[**Q&amp;A** <<x>>](#summary)", "Q&A \u00ABx\u00BB")]
    public async Task RendersInsideLinkText(string markdown, string expected)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();
        var hyperlink = paragraph.Elements<Hyperlink>().Single();

        await Assert.That(paragraph.InnerText).IsEqualTo(expected);
        await Assert.That(hyperlink.InnerText).IsEqualTo(expected);
        foreach (var run in hyperlink.Elements<Run>())
        {
            await Assert.That(run.RunProperties?.RunStyle?.Val?.Value).IsEqualTo("Hyperlink");
        }
    }

    // The link's target is decoded by Markdig, the text by the renderer; neither is decoded twice.
    [Test]
    public async Task EntityInALinkTargetIsDecodedOnce()
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(RendererHarness.FirstInline<LinkInline>("[Q&amp;A](https://example.com/?a=1&amp;b=2)"));

        await Assert.That(renderer.Top.CurrentRuns.Single().InnerText).IsEqualTo("Q&A");
        await Assert.That(renderer.MainPart.HyperlinkRelationships.Single().Uri.OriginalString)
            .IsEqualTo("https://example.com/?a=1&b=2");
    }

    // The text these produce is text like any other inside the emphasis: one run, formatted.
    [Test]
    [Arguments("**Q&amp;A [x <<y>>**", "Q&A [x \u00ABy\u00BB", "b")]
    [Arguments("*a &amp; [b*", "a & [b", "i")]
    [Arguments("~~a [b~~", "a [b", "strike")]
    [Arguments("<u>P&amp;L <<x>> [y</u>", "P&L \u00ABx\u00BB [y", "u")]
    public async Task TakesTheFormattingAroundIt(string markdown, string expected, string properties)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();
        var run = paragraph.Elements<Run>().Single();

        await Assert.That(run.InnerText).IsEqualTo(expected);
        await Assert.That(string.Join(' ', run.RunProperties!.Select(_ => _.LocalName))).IsEqualTo(properties);
    }

    // Alt text is read off the same inlines by a walk of its own, and is what shows in place of an
    // image that is not loaded - here, one the policy refuses.
    [Test]
    [Arguments("![Q&amp;A <<x>> &copy;](https://example.invalid/missing.png)", "Q&A \u00ABx\u00BB \u00A9")]
    [Arguments("![the `code` and *the | pipe*](https://example.invalid/missing.png)", "the code and the | pipe")]
    [Arguments("![two\nlines](https://example.invalid/missing.png)", "two lines")]
    [Arguments("![see <https://example.com/a>](https://example.invalid/missing.png)", "see https://example.com/a")]
    public async Task ImageAltTextKeepsThem(string markdown, string expected)
    {
        var image = RendererHarness.FirstInline<LinkInline>(markdown);
        var mainPart = RendererHarness.BuildRenderer().MainPart;
        var renderer = new OpenXmlMarkdownRenderer(
            mainPart,
            new(mainPart),
            new(OpenXmlHtml.ImagePolicy.Deny(), OpenXmlHtml.ImagePolicy.Deny()));

        renderer.Render(image);

        await Assert.That(string.Concat(renderer.Top.CurrentRuns.Select(_ => _.InnerText))).IsEqualTo(expected);
    }
}
