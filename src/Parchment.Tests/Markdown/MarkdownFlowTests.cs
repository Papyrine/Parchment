public class MarkdownFlowTests
{
    public class ReportModel
    {
        public required string Title { get; init; }
        public required string Author { get; init; }
        public required IReadOnlyList<string> Findings { get; init; }
    }

    [Test]
    public async Task BasicMarkdown()
    {
        var markdown =
            """
            # {{ Title }}

            by *{{ Author }}*

            ## Key findings

            {% for finding in Findings %}
            - {{ finding }}
            {% endfor %}

            > Review complete.
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportModel>("report", markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            "report",
            new ReportModel
            {
                Title = "Q2 Engineering Review",
                Author = "Alex Chen",
                Findings =
                [
                    "Build times improved 40%",
                    "Test flake rate halved",
                    "Three new services in production"
                ]
            },
            stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    public class TitleModel
    {
        public required string Title { get; init; }
    }

    #region MarkdownTemplatePropertyModel

    public class BriefModel
    {
        public required string Title;
        public required string Details;
    }

    #endregion

    [Test]
    public async Task PropertyContainingMarkdown()
    {
        using var targetStream = new MemoryStream();
        var markdown =
            """
            <!-- begin-snippet: MarkdownTemplatePropertyContent(lang=handlebars) -->
            # {{ Title }}

            {{ Details }}
            <!-- end-snippet -->
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        #region MarkdownTemplatePropertyUsage

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<BriefModel>(
            "brief",
            markdown,
            styleSource);

        await store.Render(
            "brief",
            new BriefModel
            {
                Title = "Sprint recap",
                Details =
                    """
                    ## Done

                    - Landed the **search** feature
                    - Fixed _three_ regressions

                    > Ship it.
                    """
            },
            targetStream);

        #endregion

        targetStream.Position = 0;
        await Verify(targetStream, "docx");
    }

    [Test]
    public async Task PropertyContainingHtml()
    {
        using var targetStream = new MemoryStream();
        var markdown =
            """
            # {{ Title }}

            {{ Details }}
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<BriefModel>(
            "brief-html",
            markdown,
            styleSource);

        await store.Render(
            "brief-html",
            new BriefModel
            {
                Title = "Release notes",
                Details =
                    """
                    <p>The <b>search</b> feature has landed.</p>
                    <ul>
                      <li>Closed three regressions</li>
                      <li>Halved test flake rate</li>
                    </ul>
                    <blockquote>Ship it.</blockquote>
                    """
            },
            targetStream);

        targetStream.Position = 0;
        await Verify(targetStream, "docx");
    }

    [Test]
    public async Task HtmlCommentsAreStripped()
    {
        // HTML comment blocks (snippet markers, authoring notes, TODOs) must not bleed into the
        // rendered docx as blank paragraphs. Two markdowns that differ only by surrounding
        // comment lines should produce byte-identical output.
        var withComments =
            """
            <!-- begin-snippet: report(lang=handlebars) -->
            # {{ Title }}

            <!-- TODO: add executive summary -->
            Body text follows the heading.
            <!-- end-snippet -->
            """;

        var withoutComments =
            """
            # {{ Title }}

            Body text follows the heading.
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TitleModel>("with-comments", withComments, styleSource);
        styleSource.Position = 0;
        store.RegisterMarkdownTemplate<TitleModel>("without-comments", withoutComments, styleSource);

        var model = new TitleModel {Title = "Sample"};

        using var withStream = new MemoryStream();
        await store.Render("with-comments", model, withStream);

        using var withoutStream = new MemoryStream();
        await store.Render("without-comments", model, withoutStream);

        await Assert.That(withStream.ToArray()).IsEquivalentTo(withoutStream.ToArray());
    }

    public class ImageModel
    {
        public required string Caption { get; init; }
    }

    [Test]
    public async Task ImageWithDataUriEmbedsDrawing()
    {
        // 1x1 transparent PNG
        const string dataUri =
            "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=";

        var markdown =
            "# {{ Caption }}\n\n![pixel](" + dataUri + ")\n";

        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ImageModel>("image", markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            "image",
            new ImageModel
            {
                Caption = "With image"
            },
            stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        var drawings = main.Document!.Body!.Descendants<Drawing>().ToList();
        await Assert.That(drawings.Count).IsEqualTo(1);
        await Assert.That(main.ImageParts.Any()).IsTrue();
    }

    static byte[] OnePixelPng() =>
        Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=");

    [Test]
    public async Task ImageFromLocalFileEmbedsDrawing()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"parchment-md-img-{Guid.NewGuid():N}.png");
        await File.WriteAllBytesAsync(pngPath, OnePixelPng());
        try
        {
            var markdown = "# {{ Caption }}\n\n![pixel](" + pngPath.Replace("\\", "/") + ")\n";

            using var styleSource = DocxTemplateBuilder.Build();
            var store = new TemplateStore();
            store.RegisterMarkdownTemplate<ImageModel>("image", markdown, styleSource);

            using var stream = new MemoryStream();
            await store.Render(
                "image",
                new ImageModel
                {
                    Caption = "With image"
                },
                stream);
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
    public async Task ImageFromLocalFileBlockedByDenyPolicy()
    {
        var pngPath = Path.Combine(Path.GetTempPath(), $"parchment-md-img-{Guid.NewGuid():N}.png");
        await File.WriteAllBytesAsync(pngPath, OnePixelPng());
        try
        {
            var markdown = "# {{ Caption }}\n\n![pixel](" + pngPath.Replace("\\", "/") + ")\n";

            using var styleSource = DocxTemplateBuilder.Build();
            var store = new TemplateStore
            {
                LocalImages = OpenXmlHtml.ImagePolicy.Deny()
            };
            store.RegisterMarkdownTemplate<ImageModel>("image", markdown, styleSource);

            using var stream = new MemoryStream();
            await store.Render(
                "image",
                new ImageModel
                {
                    Caption = "With image"
                },
                stream);
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

    public class LoopModel
    {
        public required IReadOnlyList<Row> Rows { get; init; }
    }

    public class Row
    {
        public required string Name { get; init; }
    }

    static void Register(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        new TemplateStore().RegisterMarkdownTemplate<LoopModel>("t", markdown, styleSource);
    }

    // Leading whitespace control is exactly what a markdown template needs, since markdown ends an
    // html block at the first blank line. The old text scan looked for a literal "{% for row in "
    // and so rejected this valid template.
    [Test]
    public void LeadingWhitespaceControlOnForIsAccepted() =>
        Register("{%- for row in Rows %}{{ row.Name }}{% endfor %}");

    [Test]
    public void TrailingWhitespaceControlOnForIsAccepted() =>
        Register("{% for row in Rows -%}{{ row.Name }}{% endfor %}");

    [Test]
    public void BothWhitespaceControlsOnForAreAccepted() =>
        Register("{%- for row in Rows -%}{{ row.Name }}{% endfor %}");

    // The old scan skipped the whole subtree once it decided a root was a loop variable, so this
    // typo threw nothing at registration or render — the if went false and the body vanished.
    [Test]
    public async Task TypoOnLoopVariableMemberFailsRegistration()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{% for row in Rows %}{% if row.NoSuchMember %}x{% endif %}{% endfor %}"));
        await Assert.That(exception.Message).Contains("NoSuchMember");
    }

    [Test]
    public async Task TypoOnLoopVariableSubstitutionFailsRegistration()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{%- for row in Rows %}{{ row.Nope }}{% endfor %}"));
        await Assert.That(exception.Message).Contains("Nope");
    }

    [Test]
    public async Task TypoOnLoopSourceFailsRegistration()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{% for row in NoSuchCollection %}{{ row.Name }}{% endfor %}"));
        await Assert.That(exception.Message).Contains("NoSuchCollection");
    }

    // A loop variable must not outlive its loop.
    [Test]
    public async Task LoopVariableDoesNotLeakPastItsLoop()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{% for row in Rows %}{{ row.Name }}{% endfor %}{{ row.Name }}"));
        await Assert.That(exception).IsNotNull();
    }

    [Test]
    public void NestedLoopsBindIndependently() =>
        Register("{% for row in Rows %}{% for inner in Rows %}{{ inner.Name }}{{ row.Name }}{% endfor %}{% endfor %}");

    // forloop is introduced by liquid itself and is not a model member.
    [Test]
    public void ForLoopIsAccepted() =>
        Register("{% for row in Rows %}{{ forloop.index }}{{ row.Name }}{% endfor %}");

    [Test]
    public void AssignIsAccepted() =>
        Register("{% assign total = Rows %}{% for row in total %}{{ row.Name }}{% endfor %}");

    [Test]
    public void CaptureIsAccepted() =>
        Register("{% capture heading %}Report{% endcapture %}{{ heading }}");

    // An assign whose value is a typo must still fail — the value expression is validated.
    [Test]
    public async Task AssignOfUnknownMemberFailsRegistration()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{% assign total = NoSuchThing %}{{ total }}"));
        await Assert.That(exception.Message).Contains("NoSuchThing");
    }

    // Overriding VisitForStatement means base is never called, so the else branch has to be walked
    // explicitly. Without that a typo in it registers clean and the branch renders empty.
    [Test]
    public async Task TypoInForElseBranchFailsRegistration()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{% for row in Rows %}{{ row.Name }}{% else %}{{ NoSuchThing }}{% endfor %}"));
        await Assert.That(exception.Message).Contains("NoSuchThing");
    }

    [Test]
    public void ValidForElseBranchIsAccepted() =>
        Register("{% for row in Rows %}{{ row.Name }}{% else %}{{ Rows }}{% endfor %}");

    // Looping something that resolves but is not enumerable is a mistake, not an unknown. The docx
    // validator and the source generator both reject it; markdown used to accept it as untyped and
    // render nothing.
    [Test]
    public async Task NonEnumerableLoopSourceFailsRegistration()
    {
        // Row is a POCO. A string would not do here — it is IEnumerable<char>, so it resolves.
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{% for row in Rows %}{% for inner in row %}{{ inner }}{% endfor %}{% endfor %}"));
        await Assert.That(exception.Message).Contains("enumerable");
    }

    // A range has no member path at all, so nothing about it is knowable and it stays accepted.
    [Test]
    public void RangeLoopSourceIsAccepted() =>
        Register("{% for i in (1..5) %}{{ i }}{% endfor %}");

    // forloop is scoped to the loop body, so outside one it is still an error — but the message
    // should say what forloop is rather than describing it as a missing model member.
    [Test]
    public async Task ForLoopOutsideALoopFailsWithAnExplanation()
    {
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => Register("{{ forloop.index }}"));
        await Assert.That(exception.Message).Contains("{% for %}");
        await Assert.That(exception.Message).DoesNotContain("is not a member of");
    }

    static MemoryStream BuildDotx()
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Template))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(new Paragraph()));
        }

        stream.Position = 0;
        return stream;
    }

    // The style source is cloned and the clone becomes the output, so a .dotx used to produce a
    // template-typed package that Word opened as a new unsaved document rather than the document.
    [Test]
    public async Task DotxStyleSourceProducesDocumentTypedOutput()
    {
        using var dotx = BuildDotx();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TitleModel>("t", "# {{ Title }}", dotx);

        using var output = new MemoryStream();
        await store.Render(
            "t",
            new TitleModel
            {
                Title = "x"
            },
            output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        await Assert.That(result.DocumentType).IsEqualTo(WordprocessingDocumentType.Document);
    }

    public class ItemsModel
    {
        public required IReadOnlyList<string> Items { get; init; }
    }

    public class TokenModel
    {
        public required TokenValue Value { get; init; }
    }

    static async Task<string> RenderText<TModel>(string markdown, TModel model)
        where TModel : class
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TModel>("t", markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render("t", model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }

    // These filters build an OpenXmlToken for the docx flow. The markdown flow has no OpenXML to
    // substitute into, so without a markdown-source form the token reached the writer and Fluid
    // wrote "Parchment.OpenXmlToken" into the document.
    [Test]
    public async Task BulletListFilterRendersListInMarkdownFlow()
    {
        var text = await RenderText(
            "{{ Items | bullet_list }}",
            new ItemsModel
            {
                Items = ["alpha", "beta"]
            });
        await Assert.That(text).IsEqualTo("alphabeta");
        await Assert.That(text).DoesNotContain("OpenXmlToken");
    }

    [Test]
    public async Task NumberedListFilterRendersListInMarkdownFlow()
    {
        var text = await RenderText(
            "{{ Items | numbered_list }}",
            new ItemsModel
            {
                Items = ["alpha", "beta"]
            });
        await Assert.That(text).IsEqualTo("alphabeta");
        await Assert.That(text).DoesNotContain("OpenXmlToken");
    }

    [Test]
    public async Task MarkdownFilterRendersSourceInMarkdownFlow()
    {
        var text = await RenderText(
            "{{ Items[0] | markdown }}",
            new ItemsModel
            {
                Items = ["**bold**"]
            });
        await Assert.That(text).IsEqualTo("bold");
        await Assert.That(text).DoesNotContain("MarkdownToken");
    }

    [Test]
    public async Task MarkdownTokenPropertyRendersSource()
    {
        var text = await RenderText(
            "{{ Value }}",
            new TokenModel
            {
                Value = new MarkdownToken("# Heading")
            });
        await Assert.That(text).IsEqualTo("Heading");
    }

    [Test]
    public async Task HtmlTokenPropertyRendersSource()
    {
        var text = await RenderText(
            "{{ Value }}",
            new TokenModel
            {
                Value = new HtmlToken("<p>from html</p>")
            });
        await Assert.That(text).IsEqualTo("from html");
    }

    [Test]
    public async Task PlainTextTokenPropertyRendersValue()
    {
        var text = await RenderText(
            "{{ Value }}",
            new TokenModel
            {
                Value = "just text"
            });
        await Assert.That(text).IsEqualTo("just text");
    }

    // An OpenXmlToken has no markdown form, so it has to fail loudly rather than stringify.
    [Test]
    public async Task OpenXmlTokenPropertyThrowsInMarkdownFlow()
    {
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () => await RenderText(
                "{{ Value }}",
                new TokenModel
                {
                    Value = new OpenXmlToken(_ => [])
                }));
        await Assert.That(exception!.Message).Contains("OpenXmlToken");
        await Assert.That(exception.Message).Contains("RegisterDocxTemplate");
    }
}
