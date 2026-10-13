// ReSharper disable PartialTypeWithSinglePart
public partial class FormatAttributeTests
{
    static string ScenarioPath(string scenarioName) =>
        Path.Combine(
            ProjectFiles.ProjectDirectory,
            "Scenarios",
            scenarioName);

    #region HtmlAttribute
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;
    #endregion

    #region MarkdownAttribute
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class MarkdownAttribute : Attribute;
    #endregion

    #region MarkdownFieldModel
    [ParchmentBindable]
    public partial record MarkdownFieldDoc
    {
        public required string Title;

        [Markdown]
        public required string Body;
    }
    #endregion

    #region HtmlModel
    [ParchmentBindable]
    public partial class HtmlDoc
    {
        public required string Title;

        [Html]
        public required string Body;
    }
    #endregion

    #region MarkdownModel
    [ParchmentBindable]
    public partial class MarkdownDoc
    {
        public required string Title;

        [Markdown]
        public required string Body;
    }
    #endregion

    #region StringSyntaxHtmlModel
    [ParchmentBindable]
    public partial class StringSyntaxHtmlDoc
    {
        public required string Title;

        [StringSyntax("html")]
        public required string Body;
    }
    #endregion

    #region StringSyntaxMarkdownModel
    public class StringSyntaxMarkdownDoc
    {
        public required string Title;

        [StringSyntax("markdown")]
        public required string Body;
    }
    #endregion



    [ParchmentBindable]
    public partial class TwoHtmlBlockDoc
    {
        [Html]
        public required string A { get; init; }

        [Html]
        public required string B { get; init; }
    }

    [Test]
    public async Task RenderHtml()
    {
        #region HtmlUsage
        var templatePath = Path.Combine(ScenarioPath("html-property"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlDoc>(templatePath);

        var model = new HtmlDoc
        {
            Title = "Report",
            Body = "<p>Hello <b>world</b></p><p>Second para</p>"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        #endregion

        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("html-property"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }

    [Test]
    public async Task RenderMarkdown()
    {
        #region MarkdownUsage
        var templatePath = Path.Combine(ScenarioPath("markdown-property"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownDoc>(templatePath);

        var model = new MarkdownDoc
        {
            Title = "Report",
            Body = "# Heading\n\nSome **bold** text."
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        #endregion

        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("markdown-property"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }

    [Test]
    public async Task MarkdownOnFieldRendersAsMarkdown()
    {
        // Regression: `[Markdown]` on a public field (e.g. record with `public required string X;`
        // and no `{ get; init; }`) must be detected by the reflection FormatMap walker — not just
        // properties. Without this, the value would be inserted as plain text and bare URLs /
        // mailto: prefixes would not autolink.
        using var template = DocxTemplateBuilder.Build("{{ Body }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownFieldDoc>(template);

        var model = new MarkdownFieldDoc
        {
            Title = "T",
            Body = "Contact (mailto:user@example.com) please."
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var hyperlinkRels = main.HyperlinkRelationships.ToList();
        await Assert.That(hyperlinkRels.Count).IsEqualTo(1);
        await Assert.That(hyperlinkRels[0].Uri.ToString()).IsEqualTo("mailto:user@example.com");
        await Assert.That(main.Document!.Body!.Descendants<Hyperlink>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task RenderMarkdownWithNestedHtml()
    {
        #region MarkdownWithHtmlUsage
        var templatePath = Path.Combine(ScenarioPath("markdown-with-html"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownDoc>(templatePath);

        var model = new MarkdownDoc
        {
            Title = "Report",
            Body = """
                # Heading

                Some **bold** markdown with an <em>inline HTML</em> tag.

                <div>
                  <p>This is an <strong>HTML block</strong> nested inside markdown.</p>
                </div>

                Back to *markdown* after the HTML block.
                """
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        #endregion

        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("markdown-with-html"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }

    [Test]
    public async Task StringSyntaxHtmlIsEquivalentToHtmlAttribute()
    {
        using var template = DocxTemplateBuilder.Build("{{ Body }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<StringSyntaxHtmlDoc>(template);

        var model = new StringSyntaxHtmlDoc
        {
            Title = "T",
            Body = "<p>via <i>StringSyntax</i></p>"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    public class VirtualHtmlBase
    {
        [Html]
        public virtual string Body { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class OverridingHtmlDoc : VirtualHtmlBase
    {
        public override string Body { get; set; } = "";
    }

    [Test]
    public async Task HtmlOnOverriddenMember_IsRendered()
    {
        // The override declares no attribute of its own; C# gives it the base member's [Html].
        using var template = DocxTemplateBuilder.Build("{{ Body }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<OverridingHtmlDoc>(template);

        var model = new OverridingHtmlDoc
        {
            Body = "<b>bold</b>"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(body.InnerText).IsEqualTo("bold");
        await Assert.That(body.Descendants<Bold>().Any()).IsTrue();
    }

    [Test]
    public async Task NonSoloHtml_InlineContent_SplicesIntoHostParagraph()
    {
        // Inline-only HTML in a non-solo token — the produced single paragraph is unwrapped and
        // its runs are spliced into the host paragraph, preserving the surrounding "Prefix " text.
        using var template = DocxTemplateBuilder.Build("Prefix {{ Body }} suffix");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlDoc>(template);

        var model = new HtmlDoc
        {
            Title = "T",
            Body = "<b>bold</b> and <i>italic</i>"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task SoloHtml_TrailingSpace_ReplacesHostParagraph()
    {
        // A space after the token leaves it alone in its paragraph: the blocks replace the host
        // and no empty before/after halves are left around them.
        using var template = DocxTemplateBuilder.Build("{{ Body }} ");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlDoc>(template);

        var model = new HtmlDoc
        {
            Title = "T",
            Body = "<p>First</p><p>Second</p>"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs.Select(_ => _.InnerText)).IsEquivalentTo(["First", "Second"]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_SplitsHostParagraph()
    {
        // Block-level HTML (multiple `<p>`s) in a non-solo token — the host paragraph is split:
        // "Prefix " becomes its own paragraph, the produced `<p>`s slot in between, and " suffix"
        // becomes another paragraph after.
        using var template = DocxTemplateBuilder.Build("Prefix {{ Body }} suffix");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlDoc>(template);

        var model = new HtmlDoc
        {
            Title = "T",
            Body = "<p>First</p><p>Second</p>"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task MarkdownWithNestedHtml_RendersInlineAndBlockHtmlAsRealFormatting()
    {
        // Inline HTML tags (<em>, <strong>, ...) push/pop RunProperties via HtmlInlineRenderer so
        // the enclosed text picks up Italic/Bold. Block HTML dispatches through OpenXmlHtml so
        // <strong> inside a <div> becomes a real Bold run. Neither set of tags survives as
        // literal text in the rendered docx.
        using var template = DocxTemplateBuilder.Build("{{ Body }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownDoc>(template);

        var model = new MarkdownDoc
        {
            Title = "T",
            Body = """
                Some **bold** with an <em>inline italic</em> tag.

                <div><p>Block <strong>HTML</strong> inside markdown.</p></div>

                Tail.
                """
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphTexts = body.Elements<Paragraph>()
            .Select(p => p.InnerText)
            .Where(t => t.Length > 0)
            .ToList();

        await Assert.That(paragraphTexts).Contains("Some bold with an inline italic tag.");
        await Assert.That(paragraphTexts).Contains("Block HTML inside markdown.");
        await Assert.That(paragraphTexts).Contains("Tail.");

        var bodyText = body.InnerText;
        await Assert.That(bodyText).DoesNotContain("<em>");
        await Assert.That(bodyText).DoesNotContain("</em>");
        await Assert.That(bodyText).DoesNotContain("<strong>");
        await Assert.That(bodyText).DoesNotContain("</strong>");

        var bolds = body.Descendants<Run>()
            .Where(r => r.RunProperties?.Bold is not null)
            .Select(r => r.InnerText)
            .ToList();
        await Assert.That(bolds).Contains("bold");
        await Assert.That(bolds).Contains("HTML");

        var italics = body.Descendants<Run>()
            .Where(r => r.RunProperties?.Italic is not null)
            .Select(r => r.InnerText)
            .ToList();
        await Assert.That(italics).Contains("inline italic");
    }

    [Test]
    public async Task NonSoloMarkdown_InlineContent_SplicesIntoHostParagraph()
    {
        using var template = DocxTemplateBuilder.Build("Note: {{ Body }} (end)");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownDoc>(template);

        var model = new MarkdownDoc
        {
            Title = "T",
            Body = "**bold** and *italic*"
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task NonSoloMarkdown_BlockContent_SplitsHostParagraph()
    {
        using var template = DocxTemplateBuilder.Build("Section: {{ Body }} (end)");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownDoc>(template);

        var model = new MarkdownDoc
        {
            Title = "T",
            Body = "First paragraph.\n\nSecond paragraph."
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task FilterOnFormatTokenIsRejected()
    {
        using var template = DocxTemplateBuilder.Build("{{ Body | upcase }}");

        var store = new TemplateStore();
        var exception = await Assert.That(
            () => store.RegisterDocxTemplate<HtmlDoc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("plain member-access");
    }

    [Test]
    public async Task TwoBlockShapedFormatTokensInSameParagraph_Throws()
    {
        // Two non-solo block-shaped structural tokens in one paragraph would require overlapping
        // host splits. ParagraphSplicer doesn't compose them — the second one must throw and
        // tell the author to put one on its own line.
        using var template = DocxTemplateBuilder.Build("Prefix {{ A }} between {{ B }} suffix");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<TwoHtmlBlockDoc>(template);

        var model = new TwoHtmlBlockDoc
        {
            A = "<p>A1</p><p>A2</p>",
            B = "<p>B1</p><p>B2</p>"
        };

        using var stream = new MemoryStream();
        var exception = await Assert.That(() => store.Render(model, stream))
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.Message).Contains("Move one of the tokens to its own paragraph");
    }

    // Markdown reaches the same table renderer from a docx template, so a table wider than Word's
    // 63 columns is refused here too, and the error names the token the markdown came through.
    [Test]
    [Arguments("{{ Body }}", "{{ Body }}")]
    [Arguments("Intro {{ Body }} outro", "{{ Body }}")]
    [Arguments("{{ Title | markdown }}", "{{ Title | markdown }}")]
    public async Task MarkdownTableOverTheColumnLimitIsRejected(string templateText, string token)
    {
        using var template = DocxTemplateBuilder.Build(templateText);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownDoc>(template);

        var model = new MarkdownDoc
        {
            Title = WideTable,
            Body = WideTable
        };

        using var stream = new MemoryStream();
        var exception = await Assert.That(() => store.Render(model, stream))
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.Message).Contains("64 columns");
        await Assert.That(exception.Message).Contains("at most 63");
        await Assert.That(exception.TokenSource).IsEqualTo(token);
        await Assert.That(exception.PartUri).IsEqualTo("/word/document.xml");
    }

    [ParchmentBindable]
    public partial class MarkdownItemsDoc
    {
        public required string[] Items { get; init; }
    }

    [Test]
    public async Task MarkdownTableOverTheColumnLimitInsideALoopIsRejected()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for item in Items %}

            {{ item | markdown }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<MarkdownItemsDoc>(template);

        var model = new MarkdownItemsDoc
        {
            Items = ["plain", WideTable]
        };

        using var stream = new MemoryStream();
        var exception = await Assert.That(() => store.Render(model, stream))
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.Message).Contains("64 columns");
        await Assert.That(exception.TokenSource).IsEqualTo("{{ item | markdown }}");
    }

    static string WideTable { get; } =
        string.Join(
            '\n',
            "|" + string.Concat(Enumerable.Repeat(" h |", 64)),
            "|" + string.Concat(Enumerable.Repeat("---|", 64)),
            "|" + string.Concat(Enumerable.Repeat(" c |", 64)));

    static async Task<Body> RenderBody<T>(string templateText, T model)
        where T : notnull
    {
        using var template = DocxTemplateBuilder.Build(templateText);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    static string ParagraphTexts(Body body) =>
        string.Join('|', body.Elements<Paragraph>().Select(_ => _.InnerText));

    [Test]
    public async Task NonSoloHtml_BlockContent_SubstitutesTokenBeforeIt()
    {
        // The split clones the host into its before/after halves. A token at a lower offset is
        // applied after the split (tokens run in reverse offset order), so it has to land in the
        // half that ends up in the document rather than in the discarded host.
        var body = await RenderBody(
            "Dear {{ Title }}: {{ Body }} end",
            new HtmlDoc
            {
                Title = "Alice",
                Body = "<p>one</p><p>two</p>"
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Dear Alice: |one|two| end");
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_SubstitutesTokenAfterIt()
    {
        var body = await RenderBody(
            "Start {{ Body }} from {{ Title }} end",
            new HtmlDoc
            {
                Title = "Alice",
                Body = "<p>one</p><p>two</p>"
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Start |one|two| from Alice end");
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_SubstitutesTokensOnBothSides()
    {
        var body = await RenderBody(
            "{{ Title }} before {{ Body }} after {{ Title }}",
            new HtmlDoc
            {
                Title = "Alice",
                Body = "<p>one</p><p>two</p>"
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Alice before |one|two| after Alice");
    }

    [Test]
    public async Task NonSoloHtml_Table_SubstitutesTokenBeforeIt()
    {
        var body = await RenderBody(
            "Dear {{ Title }}: {{ Body }} end",
            new HtmlDoc
            {
                Title = "Alice",
                Body = "<table><tr><td>cell</td></tr></table>"
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Dear Alice: | end");
        await Assert.That(body.Elements<Table>().Single().InnerText).IsEqualTo("cell");
        await Assert.That(body.ChildElements[1]).IsTypeOf<Table>();
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsLineBreakInTokenBeforeIt()
    {
        var body = await RenderBody(
            "Dear {{ Title }}: {{ Body }} end",
            new HtmlDoc
            {
                Title = "Alice\nSmith",
                Body = "<p>one</p><p>two</p>"
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(ParagraphTexts(body)).IsEqualTo("Dear AliceSmith: |one|two| end");
        await Assert.That(paragraphs[0].Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<Break>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task NonSoloMarkdown_BlockContent_SubstitutesTokenBeforeIt()
    {
        var body = await RenderBody(
            "Dear {{ Title }}: {{ Body }} end",
            new MarkdownDoc
            {
                Title = "Alice",
                Body = "First paragraph.\n\nSecond paragraph."
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Dear Alice: |First paragraph.|Second paragraph.| end");
    }

    [Test]
    public async Task InlineFormatTokenBeforeBlockShapedOne_SplicesIntoBeforeHalf()
    {
        var body = await RenderBody(
            "Prefix {{ A }} between {{ B }} suffix",
            new TwoHtmlBlockDoc
            {
                A = "<b>bold</b>",
                B = "<p>B1</p><p>B2</p>"
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Prefix bold between |B1|B2| suffix");
        var bolds = body.Elements<Paragraph>()
            .First()
            .Descendants<Run>()
            .Where(_ => _.RunProperties?.Bold is not null)
            .Select(_ => _.InnerText)
            .ToList();
        await Assert.That(bolds).IsEquivalentTo(["bold"]);
    }

    [Test]
    public async Task InlineFormatTokenAfterBlockShapedOne_SplicesIntoAfterHalf()
    {
        var body = await RenderBody(
            "Prefix {{ A }} between {{ B }} suffix",
            new TwoHtmlBlockDoc
            {
                A = "<p>A1</p><p>A2</p>",
                B = "<b>bold</b>"
            });

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Prefix |A1|A2| between bold suffix");
    }

    [Test]
    public async Task NonSoloHtml_InlineContent_KeepsTabRunBeforeTokenOnce()
    {
        // A run holding only a tab carries no text, so a text-driven trim keeps it in both the
        // before and the after half. It sits before the token and must land there exactly once.
        using var template = BuildSingleParagraphTemplate(
            new Run(new Text("Label")),
            new Run(new TabChar()),
            PreservedRun("{{ Body }} tail"));

        var paragraph = (await RenderBodyParagraphs(template, "<b>x</b>")).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("Labelx tail");
        await Assert.That(paragraph.Descendants<TabChar>().Count()).IsEqualTo(1);
        await Assert.That(TextBefore(paragraph, paragraph.Descendants<TabChar>().Single())).IsEqualTo("Label");
    }

    [Test]
    public async Task NonSoloHtml_InlineContent_KeepsBreakAfterTokenInSameRunOnce()
    {
        using var template = BuildSingleParagraphTemplate(
            new Run(
                new Text("pre {{ Body }} post")
                {
                    Space = SpaceProcessingModeValues.Preserve
                },
                new Break()));

        var paragraph = (await RenderBodyParagraphs(template, "<b>x</b>")).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("pre x post");
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(TextBefore(paragraph, paragraph.Descendants<Break>().Single())).IsEqualTo("pre x post");
    }

    [Test]
    public async Task NonSoloHtml_InlineContent_KeepsNonTextContentAroundTokenInSameRun()
    {
        // The token's own run carries a tab before the token text and a break after it; the run
        // has no other text, so removing the token text must not take either element with it.
        using var template = BuildSingleParagraphTemplate(
            PreservedRun("pre "),
            new Run(
                new TabChar(),
                new Text("{{ Body }}"),
                new Break()),
            PreservedRun(" post"));

        var paragraph = (await RenderBodyParagraphs(template, "<b>x</b>")).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("pre x post");
        await Assert.That(paragraph.Descendants<TabChar>().Count()).IsEqualTo(1);
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(TextBefore(paragraph, paragraph.Descendants<TabChar>().Single())).IsEqualTo("pre ");
        await Assert.That(TextBefore(paragraph, paragraph.Descendants<Break>().Single())).IsEqualTo("pre x");
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsTabRunBeforeTokenOnce()
    {
        using var template = BuildSingleParagraphTemplate(
            new Run(new Text("Label")),
            new Run(new TabChar()),
            PreservedRun("{{ Body }} tail"));

        var paragraphs = await RenderBodyParagraphs(template, "<p>one</p><p>two</p>");

        await Assert.That(string.Join('|', paragraphs.Select(_ => _.InnerText))).IsEqualTo("Label|one|two| tail");
        await Assert.That(paragraphs.Sum(_ => _.Descendants<TabChar>().Count())).IsEqualTo(1);
        await Assert.That(paragraphs[0].Descendants<TabChar>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsBreakAfterTokenInSameRunOnce()
    {
        using var template = BuildSingleParagraphTemplate(
            new Run(
                new Text("pre {{ Body }} post")
                {
                    Space = SpaceProcessingModeValues.Preserve
                },
                new Break()));

        var paragraphs = await RenderBodyParagraphs(template, "<p>one</p><p>two</p>");

        await Assert.That(string.Join('|', paragraphs.Select(_ => _.InnerText))).IsEqualTo("pre |one|two| post");
        await Assert.That(paragraphs.Sum(_ => _.Descendants<Break>().Count())).IsEqualTo(1);
        await Assert.That(paragraphs[^1].Descendants<Break>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_SubstitutesTokenBeforeItAndKeepsTabAndBreakOnce()
    {
        // The host is trimmed in place as the before half: the leading tab stays there with the
        // substituted token, and the break after the block-shaped token goes to the after half.
        using var template = BuildSingleParagraphTemplate(
            new Run(new TabChar()),
            PreservedRun("Dear {{ Title }}: {{ Body }}"),
            new Run(new Break()),
            PreservedRun(" end"));

        var paragraphs = await RenderBodyParagraphs(template, "<p>one</p><p>two</p>");

        await Assert.That(string.Join('|', paragraphs.Select(_ => _.InnerText))).IsEqualTo("Dear T: |one|two| end");
        await Assert.That(paragraphs.Sum(_ => _.Descendants<TabChar>().Count())).IsEqualTo(1);
        await Assert.That(paragraphs[0].Descendants<TabChar>().Count()).IsEqualTo(1);
        await Assert.That(paragraphs.Sum(_ => _.Descendants<Break>().Count())).IsEqualTo(1);
        await Assert.That(paragraphs[^1].Descendants<Break>().Count()).IsEqualTo(1);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsSectionBreakOnceOnTrailingHalf()
    {
        // Word puts a section break on the last paragraph of the section it closes. Both halves
        // of a split start from the host's paragraph properties, so the break has to come off
        // the first half or the section ends twice.
        using var template = BuildBodyTemplate(
            new Paragraph(PreservedRun("section one")),
            SectionBreakParagraph("before {{ Body }} after"),
            new Paragraph(PreservedRun("section two")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");

        await Assert.That(ParagraphTexts(body)).IsEqualTo("section one|before |one|two| after|section two");
        await Assert.That(SectionBreakPositions(body)).IsEquivalentTo([4]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsSectionBreakOnEmptyTrailingHalf()
    {
        using var template = BuildBodyTemplate(
            new Paragraph(PreservedRun("section one")),
            SectionBreakParagraph("before {{ Body }}"),
            new Paragraph(PreservedRun("section two")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");

        await Assert.That(ParagraphTexts(body)).IsEqualTo("section one|before |one|two||section two");
        await Assert.That(SectionBreakPositions(body)).IsEquivalentTo([4]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsSectionBreakWhenLeadingHalfIsEmpty()
    {
        using var template = BuildBodyTemplate(
            new Paragraph(PreservedRun("section one")),
            SectionBreakParagraph("{{ Body }} after"),
            new Paragraph(PreservedRun("section two")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");

        await Assert.That(ParagraphTexts(body)).IsEqualTo("section one||one|two| after|section two");
        await Assert.That(SectionBreakPositions(body)).IsEquivalentTo([4]);
    }

    [Test]
    public async Task NonSoloHtml_Table_KeepsSectionBreakOnParagraphAfterTable()
    {
        // A section break needs a paragraph to sit on, so content that ends with a table has to
        // be followed by the trailing half even though nothing is left to put in it.
        using var template = BuildBodyTemplate(
            new Paragraph(PreservedRun("section one")),
            SectionBreakParagraph("before {{ Body }}"),
            new Paragraph(PreservedRun("section two")));

        var body = await RenderValidBody(template, "<table><tr><td>cell</td></tr></table>");

        await Assert.That(ParagraphTexts(body)).IsEqualTo("section one|before ||section two");
        await Assert.That(body.ChildElements[2]).IsTypeOf<Table>();
        await Assert.That(SectionBreakPositions(body)).IsEquivalentTo([3]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_SubstitutesTokenBeforeItInSectionBreakParagraph()
    {
        // The token at the lower offset is substituted into the host after the split has taken
        // the section break off it.
        using var template = BuildBodyTemplate(
            SectionBreakParagraph("Dear {{ Title }}: {{ Body }} end"),
            new Paragraph(PreservedRun("section two")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");

        await Assert.That(ParagraphTexts(body)).IsEqualTo("Dear T: |one|two| end|section two");
        await Assert.That(SectionBreakPositions(body)).IsEquivalentTo([3]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsPageBreakBeforeOnFirstHalfOnly()
    {
        using var template = BuildBodyTemplate(
            new Paragraph(PreservedRun("first page")),
            new Paragraph(
                new ParagraphProperties(
                    new PageBreakBefore(),
                    new Justification
                    {
                        Val = JustificationValues.Center
                    }),
                PreservedRun("before {{ Body }} after")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");
        var paragraphs = body.Elements<Paragraph>().ToList();

        await Assert.That(ParagraphTexts(body)).IsEqualTo("first page|before |one|two| after");
        await Assert.That(paragraphs.Where(_ => _.ParagraphProperties?.PageBreakBefore != null).Select(_ => _.InnerText))
            .IsEquivalentTo(["before "]);
        // Formatting still reaches both halves.
        await Assert.That(paragraphs.Where(_ => _.ParagraphProperties?.Justification != null).Select(_ => _.InnerText))
            .IsEquivalentTo(["before ", " after"]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_DropsPropertiesLeftEmptyOnTrailingHalf()
    {
        using var template = BuildBodyTemplate(
            new Paragraph(
                new ParagraphProperties(new PageBreakBefore()),
                PreservedRun("before {{ Body }} after")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");
        var paragraphs = body.Elements<Paragraph>().ToList();

        await Assert.That(paragraphs[0].ParagraphProperties?.PageBreakBefore).IsNotNull();
        await Assert.That(paragraphs[^1].ParagraphProperties).IsNull();
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_KeepsSuppressedPageBreakBeforeOnBothHalves()
    {
        // Written as off, the element is there to cancel a break the paragraph style asks for,
        // and the trailing half has the same style.
        using var template = BuildBodyTemplate(
            new Paragraph(
                new ParagraphProperties(
                    new PageBreakBefore
                    {
                        Val = false
                    }),
                PreservedRun("before {{ Body }} after")));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");

        await Assert.That(body.Elements<Paragraph>().Where(_ => _.ParagraphProperties?.PageBreakBefore != null).Select(_ => _.InnerText))
            .IsEquivalentTo(["before ", " after"]);
    }

    [Test]
    public async Task NonSoloHtml_BlockContent_NumbersSplitListItemOnce()
    {
        // One list item in the template is one number in the output: a trailing half that kept
        // the numbering would be counted as an item of its own and push every later number up.
        using var template = BuildBodyTemplate(
            ListItem("first"),
            ListItem("before {{ Body }} after"),
            ListItem("last"));

        var body = await RenderValidBody(template, "<p>one</p><p>two</p>");
        var paragraphs = body.Elements<Paragraph>().ToList();

        await Assert.That(ParagraphTexts(body)).IsEqualTo("first|before |one|two| after|last");
        await Assert.That(paragraphs.Where(IsNumbered).Select(_ => _.InnerText))
            .IsEquivalentTo(["first", "before ", "last"]);

        // Number 0 switches numbering off, including any the paragraph style would supply; the
        // rest of the item's formatting stays.
        var trailing = paragraphs[4].ParagraphProperties!;
        await Assert.That(trailing.NumberingProperties!.NumberingId!.Val!.Value).IsEqualTo(0);
        await Assert.That(trailing.ParagraphStyleId!.Val!.Value).IsEqualTo("ListParagraph");
    }

    static bool IsNumbered(Paragraph paragraph) =>
        paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value is > 0;

    static Paragraph ListItem(string text) =>
        new(
            new ParagraphProperties(
                new ParagraphStyleId
                {
                    Val = "ListParagraph"
                },
                new NumberingProperties(
                    new NumberingLevelReference
                    {
                        Val = 0
                    },
                    new NumberingId
                    {
                        Val = 1
                    })),
            PreservedRun(text));

    static Paragraph SectionBreakParagraph(string text) =>
        new(
            new ParagraphProperties(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    })),
            PreservedRun(text));

    /// <summary>
    /// Indexes, among the body's children, of the paragraphs that carry a section break.
    /// </summary>
    static List<int> SectionBreakPositions(Body body) =>
        body.ChildElements
            .Select((element, index) => (element, index))
            .Where(_ => _.element is Paragraph { ParagraphProperties.SectionProperties: not null })
            .Select(_ => _.index)
            .ToList();

    static async Task<Body> RenderValidBody(MemoryStream template, string html)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlDoc>(template);

        var model = new HtmlDoc
        {
            Title = "T",
            Body = html
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        // Named in full: a using directive would move every readme snippet this file supplies.
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();

        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    static Run PreservedRun(string text) =>
        new(
            new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

    static string TextBefore(Paragraph paragraph, OpenXmlElement element) =>
        string.Concat(
            paragraph
                .Descendants()
                .TakeWhile(_ => _ != element)
                .OfType<Text>()
                .Select(_ => _.Text));

    static async Task<List<Paragraph>> RenderBodyParagraphs(MemoryStream template, string html)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<HtmlDoc>(template);

        var model = new HtmlDoc
        {
            Title = "T",
            Body = html
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Select(_ => (Paragraph)_.CloneNode(true))
            .ToList();
    }

    static MemoryStream BuildSingleParagraphTemplate(params OpenXmlElement[] children) =>
        BuildBodyTemplate(new Paragraph(children));

    static MemoryStream BuildBodyTemplate(params OpenXmlElement[] blocks)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var body = new Body(blocks);
            body.AppendChild(
                new SectionProperties(
                    new PageSize
                    {
                        Width = 6500,
                        Height = 8000
                    }));

            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(body);
        }

        stream.Position = 0;
        return stream;
    }

    static byte[] OnePixelPng() =>
        Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=");

    [Test]
    public async Task HtmlImageFromLocalFileEmbedsDrawing()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"parchment-html-img-{Guid.NewGuid():N}.png");
        await File.WriteAllBytesAsync(pngPath, OnePixelPng());
        try
        {
            using var template = DocxTemplateBuilder.Build("{{ Body }}");
            var store = new TemplateStore();
            store.RegisterDocxTemplate<HtmlDoc>(template);

            var model = new HtmlDoc
            {
                Title = "Report",
                Body = $"<p><img src=\"{pngPath.Replace("\\", "/")}\" alt=\"pixel\" /></p>"
            };

            using var stream = new MemoryStream();
            await store.Render(model, stream);
            stream.Position = 0;

            using var doc = WordprocessingDocument.Open(stream, false);
            var main = doc.MainDocumentPart!;
            await Assert.That(main.Document!.Body!.Descendants<Drawing>().Count()).IsEqualTo(1);
            await Assert.That(main.ImageParts.Any()).IsTrue();
        }
        finally
        {
            File.Delete(pngPath);
        }
    }

    [Test]
    public async Task HtmlImageFromLocalFileBlockedByDenyPolicy()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"parchment-html-img-{Guid.NewGuid():N}.png");
        await File.WriteAllBytesAsync(pngPath, OnePixelPng());
        try
        {
            using var template = DocxTemplateBuilder.Build("{{ Body }}");
            var store = new TemplateStore
            {
                LocalImages = OpenXmlHtml.ImagePolicy.Deny()
            };
            store.RegisterDocxTemplate<HtmlDoc>(template);

            var model = new HtmlDoc
            {
                Title = "Report",
                Body = $"<p><img src=\"{pngPath.Replace("\\", "/")}\" alt=\"pixel\" /></p>"
            };

            using var stream = new MemoryStream();
            await store.Render(model, stream);
            stream.Position = 0;

            using var doc = WordprocessingDocument.Open(stream, false);
            var main = doc.MainDocumentPart!;
            await Assert.That(main.Document!.Body!.Descendants<Drawing>().Any()).IsFalse();
            await Assert.That(main.ImageParts.Any()).IsFalse();
        }
        finally
        {
            File.Delete(pngPath);
        }
    }

    [ParchmentBindable]
    public partial class NonStringFormatDoc
    {
        public int Count;
        public decimal? Price;
        public string? Note;
    }

    // The generator refuses a marker on a non-string (PARCH029), so the entries are registered by
    // hand here: a value that is not a string must print as an unmarked member does, not vanish.
    [Test]
    public async Task NonStringFormatEntryRendersTheValue()
    {
        Parchment.Generated.GeneratedRegistration.RegisterFormat(
            typeof(NonStringFormatDoc),
            [
                new("Count", Parchment.Generated.FormatMapKind.Html, _ => ((NonStringFormatDoc)_).Count),
                new("Price", Parchment.Generated.FormatMapKind.Markdown, _ => ((NonStringFormatDoc)_).Price),
                new("Note", Parchment.Generated.FormatMapKind.Html, _ => ((NonStringFormatDoc)_).Note)
            ]);

        using var template = DocxTemplateBuilder.Build("Count={{ Count }}|\nPrice={{ Price }}|\nNote={{ Note }}|");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<NonStringFormatDoc>(template);

        using var stream = new MemoryStream();
        await store.Render(
            new NonStringFormatDoc
            {
                Count = 5,
                Price = 9.5m
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var text = string.Join('\n', doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Select(_ => _.InnerText));
        await Assert.That(text).IsEqualTo("Count=5|\nPrice=9.5|\nNote=|");
    }
}
