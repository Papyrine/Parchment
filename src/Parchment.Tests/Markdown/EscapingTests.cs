// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A bound value is data, not markdown source. These pin what MarkdownEncoder guarantees, what opts
/// out of it, and that a line break in a value survives into the document as a line break.
/// </summary>
public partial class EscapingTests
{
    [ParchmentBindable]
    public partial class EscapeModel
    {
        public string Title { get; init; } = "";
        public string Details { get; init; } = "";
        public IReadOnlyList<string> Items { get; init; } = [];
        public TokenValue Token { get; init; } = "";
    }

    #region EscapingModel

    [ParchmentBindable]
    public partial class FeedbackModel
    {
        public required string Author;

        // Whatever the reviewer typed. Text, whether or not it happens to read as markdown.
        public required string Comment;

        // Written by the template's authors rather than by a reviewer, so it is markup.
        [Markdown]
        public required string Verdict;
    }

    #endregion

    // The worked example the readme shows a page render of.
    [Test]
    public async Task Documented()
    {
        var template =
            """
            <!-- begin-snippet: EscapingTemplate(lang=handlebars) -->
            # Feedback

            | Reviewer | Comment |
            | --- | --- |
            | {{ Author }} | {{ Comment }} |

            {{ Verdict }}
            <!-- end-snippet -->
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<FeedbackModel>(template, styleSource);

        using var stream = new MemoryStream();

        #region EscapingUsage

        await store.Render(
            new FeedbackModel
            {
                Author = "A. Okafor",
                Comment = "Ship it | but **check** the totals\nand the delivery dates",
                Verdict = "**Approved** with conditions"
            },
            stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    // Each of the remaining opt-outs, rendered for the readme. All follow one shape: the same value
    // bound twice, once plain and once through the opt-out, so the picture shows what the opt-out
    // changes rather than only what it produces.

    [Test]
    public async Task DocumentedMarkdownFilter()
    {
        var body = await RenderDocumented(
            """
            <!-- begin-snippet: MarkdownFilterComparison(lang=handlebars) -->
            {{ Details }}

            {{ Details | markdown }}
            <!-- end-snippet -->
            """,
            new()
            {
                Details = "**Bold** and _italic_"
            });

        await Verify(body, "docx");
    }

    [Test]
    public async Task DocumentedHtmlFilter()
    {
        var body = await RenderDocumented(
            """
            <!-- begin-snippet: HtmlFilterComparison(lang=handlebars) -->
            {{ Details }}

            {{ Details | html }}
            <!-- end-snippet -->
            """,
            new()
            {
                Details = "<b>Bold</b> and <i>italic</i>"
            });

        await Verify(body, "docx");
    }

    [Test]
    public async Task DocumentedRawFilter()
    {
        var body = await RenderDocumented(
            """
            <!-- begin-snippet: RawFilterComparison(lang=handlebars) -->
            {{ Details }}

            {{ Details | raw }}
            <!-- end-snippet -->
            """,
            new()
            {
                Details = "**Bold** and <i>italic</i>"
            });

        await Verify(body, "docx");
    }

    [Test]
    public async Task DocumentedTokenValue()
    {
        var body = await RenderDocumented(
            """
            <!-- begin-snippet: TokenValueComparison(lang=handlebars) -->
            {{ Details }}

            {{ Token }}
            <!-- end-snippet -->
            """,
            new()
            {
                Details = "**Bound as a string**",
                Token = new MarkdownToken("**Returned as a MarkdownToken**")
            });

        await Verify(body, "docx");
    }

    [Test]
    public async Task DocumentedListFilters()
    {
        var body = await RenderDocumented(
            """
            <!-- begin-snippet: ListFilterComparison(lang=handlebars) -->
            {{ Items | bullet_list }}

            {{ Items | numbered_list }}
            <!-- end-snippet -->
            """,
            new()
            {
                Items = ["Totals | and dates", "1. Not a nested list", "**Not bold**"]
            });

        await Verify(body, "docx");
    }

    static async Task<MemoryStream> RenderDocumented(string template, EscapeModel model)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EscapeModel>(template, styleSource);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    // Matched by name, the way a consumer declares them — see readme "Html and Markdown properties".
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;

    [ParchmentBindable]
    public partial class AnnotatedModel
    {
        public string Plain { get; init; } = "";

        [Markdown]
        public string Marked { get; init; } = "";

        [Html]
        public string Tagged { get; init; } = "";

        [StringSyntax("markdown")]
        public string Syntaxed { get; init; } = "";
    }

    // The original breakage: a value carrying a pipe took over the row it sat in, silently turning
    // one cell into two and shifting every column after it.
    [Test]
    public async Task PipeStaysInsideItsCell()
    {
        var body = await Render(
            """
            | Name | Note |
            | --- | --- |
            | {{ Title }} | ok |
            """,
            new()
            {
                Title = "a|b"
            });

        var cells = body.Descendants<TableCell>().ToList();
        await Assert.That(cells.Count).IsEqualTo(4);
        await Assert.That(cells[2].InnerText).IsEqualTo("a|b");
        await Assert.That(cells[3].InnerText).IsEqualTo("ok");
    }

    // Generic attributes are enabled, so an unescaped value could reach out and restyle the
    // paragraph it was substituted into — the model deciding how the document looks.
    [Test]
    public async Task StyleAttributeInAValueDoesNotRestyle()
    {
        var body = await Render(
            "{{ Title }}",
            new()
            {
                Title = "Total {.Heading1}"
            });

        var paragraph = body.Elements<Paragraph>().First();
        await Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value).IsNull();
        await Assert.That(paragraph.InnerText).IsEqualTo("Total {.Heading1}");
    }

    [Test]
    public async Task MarkdownSyntaxInAValuePrintsAsItself()
    {
        var body = await Render(
            "{{ Details }}",
            new()
            {
                Details = "## Not a heading, **not bold**, <b>not html</b>"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("## Not a heading, **not bold**, <b>not html</b>");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsFalse();
    }

    [Test]
    public async Task NewlineBecomesABreak()
    {
        var body = await Render(
            "{{ Title }}",
            new()
            {
                Title = "one\ntwo"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(paragraph.InnerText).IsEqualTo("onetwo");
    }

    // Windows line endings are one break, not two — the pair has to be seen as a pair, which is why
    // the encoder cannot lean on TextEncoder's per-character path.
    [Test]
    public async Task CarriageReturnLineFeedIsOneBreak()
    {
        var body = await Render(
            "{{ Title }}",
            new()
            {
                Title = "one\r\ntwo"
            });

        await Assert.That(body.Descendants<Break>().Count()).IsEqualTo(1);
    }

    // A blank line in a value stays inside the host paragraph rather than starting a new one, so it
    // cannot escape the paragraph's style or orphan a {.Style} attached to it.
    [Test]
    public async Task BlankLineInAValueDoesNotSplitTheParagraph()
    {
        var body = await Render(
            "{{ Title }}{.Quote}",
            new()
            {
                Title = "one\n\ntwo"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Quote");
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task RawOptsOut()
    {
        var body = await Render(
            "{{ Details | raw }}",
            new()
            {
                Details = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    [Test]
    public async Task MarkdownFilterOptsOut()
    {
        var body = await Render(
            "{{ Details | markdown }}",
            new()
            {
                Details = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    // The token type is the declaration that a value is markup, so it is what turns escaping off.
    [Test]
    public async Task MarkdownTokenOptsOut()
    {
        var body = await Render(
            "{{ Token }}",
            new()
            {
                Token = new MarkdownToken("## Heading")
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    // ...and a plain string assigned to the same member is still a plain string.
    [Test]
    public async Task StringAssignedToATokenValueIsStillEscaped()
    {
        var body = await Render(
            "{{ Token }}",
            new()
            {
                Token = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId).IsNull();
        await Assert.That(paragraph.InnerText).IsEqualTo("## Heading");
    }

    // The filter's own markers stay syntax while what came from the model does not, so an item that
    // reads like a list marker is an item rather than a nested list.
    [Test]
    public async Task ListFilterEscapesItemsButKeepsItsMarkers()
    {
        var body = await Render(
            "{{ Items | bullet_list }}",
            new()
            {
                Items = ["a|b", "1. not nested"]
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(2);
        await Assert.That(paragraphs[0].InnerText).IsEqualTo("a|b");
        await Assert.That(paragraphs[1].InnerText).IsEqualTo("1. not nested");
        await Assert.That(paragraphs[0].ParagraphProperties!.NumberingProperties).IsNotNull();
    }

    // Encoding is an output concern. Everything upstream of the writer — conditions, filter inputs,
    // loop sources — sees exactly what the model holds.
    [Test]
    public async Task ComparisonsSeeTheUnescapedValue()
    {
        var body = await Render(
            """{% if Title == "a|b" %}matched{% endif %}""",
            new()
            {
                Title = "a|b"
            });

        await Assert.That(body.InnerText).IsEqualTo("matched");
    }

    // The regression guard for the whole marker mechanism. Written straight into the source this
    // value gets markdown's answer — the span printed as literal text and *b* turned italic — so
    // an assertion that it gets html's answer instead is what proves the value reached the html
    // converter rather than Markdig.
    [Test]
    public async Task HtmlFilterAppliesHtmlSemanticsRatherThanMarkdown()
    {
        var body = await Render(
            "{{ Details | html }}",
            new()
            {
                Details = "<span>a *b* c</span>"
            });

        await Assert.That(body.InnerText).IsEqualTo("a *b* c");
        await Assert.That(body.Descendants<Italic>().Any()).IsFalse();
    }

    // Converted html is whole blocks, so it replaces the block it lands in and anything sharing
    // that block would be discarded. Refused rather than silently dropped.
    [Test]
    public async Task HtmlFilterMustBeAloneInItsBlock() =>
        await Assert.That(
                async () => await Render(
                    "Before {{ Details | html }} after",
                    new()
                    {
                        Details = "<p>Body</p>"
                    }))
            .Throws<ParchmentRenderException>();

    // An HtmlToken has to reach the converter the same way the filter's value does. It used to pass
    // through here — the one route of the three that did not convert — so this is the same guard
    // pointed at the type rather than at the filter.
    [Test]
    public async Task HtmlTokenAppliesHtmlSemanticsRatherThanMarkdown()
    {
        var body = await Render(
            "{{ Token }}",
            new()
            {
                Token = new HtmlToken("<span>a *b* c</span>")
            });

        await Assert.That(body.InnerText).IsEqualTo("a *b* c");
        await Assert.That(body.Descendants<Italic>().Any()).IsFalse();
    }

    [Test]
    public async Task HtmlTokenMustBeAloneInItsBlock() =>
        await Assert.That(
                async () => await Render(
                    "Before {{ Token }} after",
                    new()
                    {
                        Token = new HtmlToken("<p>Body</p>")
                    }))
            .Throws<ParchmentRenderException>();

    // Renders in flight at once must not see each other's parked html. The scope is async-local
    // precisely so a render owns its flow rather than its thread.
    [Test]
    public async Task ConcurrentRendersDoNotShareParkedHtml()
    {
        var bodies = await Task.WhenAll(
            Enumerable.Range(0, 8)
                .Select(index => Render(
                    "{{ Token }}",
                    new()
                    {
                        Token = new HtmlToken($"<p>Body {index}</p>")
                    })));

        for (var index = 0; index < bodies.Length; index++)
        {
            await Assert.That(bodies[index].InnerText).IsEqualTo($"Body {index}");
        }
    }

    [Test]
    public async Task HtmlFilterOptsOut()
    {
        var body = await Render(
            "{{ Details | html }}",
            new()
            {
                Details = "<p>The <b>search</b> feature has landed.</p>"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("The search feature has landed.");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsTrue();
    }

    // The annotation says once, at the model, what the filter says at every site. It reached the
    // docx flow already; a markdown template ignored it, which stopped being invisible the moment
    // an unannotated value started being escaped.
    [Test]
    public async Task MarkdownAttributeIsHonoured()
    {
        var body = await RenderAnnotated(
            "{{ Marked }}",
            new()
            {
                Marked = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    [Test]
    public async Task HtmlAttributeIsHonoured()
    {
        var body = await RenderAnnotated(
            "{{ Tagged }}",
            new()
            {
                Tagged = "<p>Plain <b>bold</b></p>"
            });

        await Assert.That(body.Descendants<Bold>().Any()).IsTrue();
        await Assert.That(body.InnerText).IsEqualTo("Plain bold");
    }

    [Test]
    public async Task StringSyntaxIsHonoured()
    {
        var body = await RenderAnnotated(
            "{{ Syntaxed }}",
            new()
            {
                Syntaxed = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    // Only the annotated member opts out. Its unannotated neighbour on the same model is text.
    [Test]
    public async Task AnnotationDoesNotLeakToOtherMembers()
    {
        var body = await RenderAnnotated(
            """
            {{ Marked }}

            {{ Plain }}
            """,
            new()
            {
                Marked = "## Rendered",
                Plain = "## Not rendered"
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs[0].ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
        await Assert.That(paragraphs[1].ParagraphProperties?.ParagraphStyleId).IsNull();
        await Assert.That(paragraphs[1].InnerText).IsEqualTo("## Not rendered");
    }

    // A filter makes the value the filter's output rather than the member, so the annotation no
    // longer describes it — the same reason the docx flow rejects that shape as PARCH010.
    [Test]
    public async Task AnnotatedMemberWithAFilterIsStillEscaped()
    {
        var body = await RenderAnnotated(
            "{{ Marked | upcase }}",
            new()
            {
                Marked = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId).IsNull();
        await Assert.That(paragraph.InnerText).IsEqualTo("## HEADING");
    }

    // A bound value can land inside the template's own link syntax, where the escaping shows up in
    // the destination. CommonMark processes backslash escapes there too, so the anchor still
    // resolves to the bookmark it names.
    [Test]
    public async Task BoundAnchorStillResolves()
    {
        var body = await Render(
            """
            # Summary {#the-section}

            [{{ Title }}](#{{ Details }})
            """,
            new()
            {
                Title = "Go to the summary",
                Details = "the-section"
            });

        var heading = body.Elements<Paragraph>().First();
        var hyperlink = body.Descendants<Hyperlink>().Single();

        await Assert.That(hyperlink.Anchor?.Value).IsEqualTo(heading.GetFirstChild<BookmarkStart>()!.Name?.Value);
        await Assert.That(hyperlink.InnerText).IsEqualTo("Go to the summary");
    }

    [Test]
    public async Task PlainTextIsUntouched()
    {
        var body = await Render(
            "{{ Title }}",
            new()
            {
                Title = "Nothing here needs escaping"
            });

        await Assert.That(body.InnerText).IsEqualTo("Nothing here needs escaping");
    }

    // The docx flow reaches the same w:br by a different route: it has a real w:t to split, so the
    // newline is parked as a marker during substitution and materialized afterwards.
    [Test]
    public async Task DocxSubstitutionNewlineBecomesABreak()
    {
        var body = await RenderDocx(
            "Before {{ Title }} after",
            new()
            {
                Title = "one\ntwo"
            });

        var paragraph = body.Elements<Paragraph>().First();
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(paragraph.InnerText).IsEqualTo("Before onetwo after");
    }

    [Test]
    public async Task DocxSubstitutionCarriageReturnLineFeedIsOneBreak()
    {
        var body = await RenderDocx(
            "{{ Title }}",
            new()
            {
                Title = "one\r\ntwo"
            });

        await Assert.That(body.Descendants<Break>().Count()).IsEqualTo(1);
    }

    // The docx flow substitutes into runs rather than into source, so it never needed escaping and
    // does not get any — markdown syntax in a value there was always just text.
    [Test]
    public async Task DocxSubstitutionIsNotEscaped()
    {
        var body = await RenderDocx(
            "{{ Details }}",
            new()
            {
                Details = "a|b **not bold**"
            });

        await Assert.That(body.InnerText).IsEqualTo("a|b **not bold**");
    }

    [Test]
    public async Task EncoderEscapesEveryStructuralCharacter()
    {
        var encoded = MarkdownEncoder.Default.Encode(@"\`*_~^[]()<>&#|{}!+-.=");

        await Assert.That(encoded).IsEqualTo(@"\\\`\*\_\~\^\[\]\(\)\<\>\&\#\|\{\}\!\+\-\.\=");
    }

    [Test]
    public async Task EncoderLeavesNonSyntaxAlone()
    {
        var value = "Plain text, with 'quotes', \"doubles\", 100% and a: colon;";

        await Assert.That(MarkdownEncoder.Default.Encode(value)).IsEqualTo(value);
    }

    [Test]
    [Arguments("a\nb", "a<br />b")]
    [Arguments("a\r\nb", "a<br />b")]
    [Arguments("a\rb", "a<br />b")]
    [Arguments("a\n\nb", "a<br /><br />b")]
    public async Task EncoderTurnsLineEndingsIntoBreaks(string value, string expected) =>
        await Assert.That(MarkdownEncoder.Default.Encode(value)).IsEqualTo(expected);

    static async Task<Body> Render(string template, EscapeModel model)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EscapeModel>(template, styleSource);
        return await RenderBody(store, model);
    }

    static async Task<Body> RenderAnnotated(string template, AnnotatedModel model)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<AnnotatedModel>(template, styleSource);

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        return (Body) document.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    static async Task<Body> RenderDocx(string template, EscapeModel model)
    {
        using var templateStream = DocxTemplateBuilder.Build(template);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EscapeModel>(templateStream);
        return await RenderBody(store, model);
    }

    static async Task<Body> RenderBody(TemplateStore store, EscapeModel model)
    {
        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        // Cloned so it outlives the package.
        return (Body) document.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }
}
