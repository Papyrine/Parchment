// ReSharper disable PartialTypeWithSinglePart
public partial class MarkdownFlowTests
{
    [ParchmentBindable]
    public partial class ReportModel
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
        store.RegisterMarkdownTemplate<ReportModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
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

    // Registering without a style source falls back to the built-in blank docx. Every other
    // markdown test supplies one, so this is the only place the built-in package's own styles —
    // and the part scan over a package that has no non-body parts — are pinned.
    [Test]
    public async Task MarkdownWithNoStyleSource()
    {
        var markdown =
            """
            # {{ Title }}

            by *{{ Author }}*

            ## Key findings

            {% for finding in Findings %}
            - {{ finding }}
            {% endfor %}

            > The quarter closed ahead of plan.

            ### Level three

            #### Level four

            ##### Level five

            ###### Level six

            | Area | Status |
            | --- | --- |
            | Build | Green |
            | Tests | Green |

            > Review complete.
            """;

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportModel>(markdown);

        using var stream = new MemoryStream();
        await store.Render(
            new ReportModel
            {
                Title = "Q2 Engineering Review",
                Author = "Alex Chen",
                Findings =
                [
                    "Build times improved 40%",
                    "Test flake rate halved"
                ]
            },
            stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class TitleModel
    {
        public required string Title { get; init; }
    }

    #region MarkdownTemplatePropertyModel

    [ParchmentBindable]
    public partial class BriefModel
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

            {{ Details | markdown }}
            <!-- end-snippet -->
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        #region MarkdownTemplatePropertyUsage

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<BriefModel>(
            markdown,
            styleSource);

        await store.Render(
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

            {{ Details | raw }}
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<BriefModel>(
            markdown,
            styleSource);

        await store.Render(
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

    // The counterpart to the two above: without an opt-out the same values are the text they are,
    // so markdown syntax and html tags print instead of restructuring the document around them.
    [Test]
    public async Task PropertyContentIsEscapedByDefault()
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
            markdown,
            styleSource);

        await store.Render(
            new BriefModel
            {
                Title = "Release notes",
                Details = "## Not a heading, <b>not bold</b>, and | not a new cell",
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
        store.RegisterMarkdownTemplate<TitleModel>(withComments, styleSource);
        styleSource.Position = 0;
        store.RegisterMarkdownTemplate<TitleModel>(withoutComments, styleSource);

        var model = new TitleModel {Title = "Sample"};

        using var withStream = new MemoryStream();
        await store.Render(model, withStream);

        using var withoutStream = new MemoryStream();
        await store.Render(model, withoutStream);

        await Assert.That(withStream.ToArray()).IsEquivalentTo(withoutStream.ToArray());
    }

    [ParchmentBindable]
    public partial class ImageModel
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
        store.RegisterMarkdownTemplate<ImageModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
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
            store.RegisterMarkdownTemplate<ImageModel>(markdown, styleSource);

            using var stream = new MemoryStream();
            await store.Render(
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
            store.RegisterMarkdownTemplate<ImageModel>(markdown, styleSource);

            using var stream = new MemoryStream();
            await store.Render(
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

    [ParchmentBindable]
    public partial class LoopModel
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
        new TemplateStore().RegisterMarkdownTemplate<LoopModel>(markdown, styleSource);
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
        store.RegisterMarkdownTemplate<TitleModel>("# {{ Title }}", dotx);

        using var output = new MemoryStream();
        await store.Render(
            new TitleModel
            {
                Title = "x"
            },
            output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        await Assert.That(result.DocumentType).IsEqualTo(WordprocessingDocumentType.Document);
    }

    [ParchmentBindable]
public partial class ItemsModel
    {
        public required IReadOnlyList<string> Items { get; init; }
    }

    [ParchmentBindable]
public partial class TokenModel
    {
        public required TokenValue Value { get; init; }
    }

    static async Task<string> RenderText<TModel>(string markdown, TModel model)
        where TModel : class
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }

    static async Task<MemoryStream> Render<TModel>(string markdown, TModel model)
        where TModel : class
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<TModel>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
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

    // An OpenXmlToken has no markdown form at all, so it is parked against a marker and the
    // marker's paragraph is swapped for what the delegate emits once Markdig has parsed - the same
    // route html and [ExcelsiorTable] members take.
    [Test]
    public async Task OpenXmlTokenPropertyRendersElements()
    {
        var text = await RenderText(
            "{{ Value }}",
            new TokenModel
            {
                Value = new OpenXmlToken(_ => [new Paragraph(new Run(new Text("from openxml")))])
            });
        await Assert.That(text).IsEqualTo("from openxml");
    }

    // The point of the token over html: bytes reach an ImagePart without a base64 detour.
    [Test]
    public async Task OpenXmlTokenPropertyCanAddAnImagePart()
    {
        using var stream = await Render(
            "{{ Value }}",
            new TokenModel
            {
                Value = new OpenXmlToken(
                    context =>
                    {
                        context.AddImagePart(OnePixelPng(), "image/png");
                        return [new Paragraph(new Run(new Text("with image")))];
                    })
            });

        using var doc = WordprocessingDocument.Open(stream, false);
        await Assert.That(doc.MainDocumentPart!.ImageParts.Count()).IsEqualTo(1);
    }

    // The delegate is handed the paragraph it is replacing, so a token that wants to inherit what
    // it is standing in for gets the same answer here as it does in the docx flow. Markdig built
    // this one rather than a template, which is the only difference.
    [Test]
    public async Task OpenXmlTokenPropertyIsGivenItsHostParagraph()
    {
        Paragraph? host = null;
        await RenderText(
            "{{ Value }}",
            new TokenModel
            {
                Value = new OpenXmlToken(
                    context =>
                    {
                        host = context.HostParagraph;
                        return [];
                    })
            });
        await Assert.That(host).IsNotNull();
    }

    [Test]
    public async Task OpenXmlTokenPropertyRenderingNothingLeavesNothing()
    {
        var text = await RenderText(
            """
            before

            {{ Value }}

            after
            """,
            new TokenModel
            {
                Value = new OpenXmlToken(_ => [])
            });
        await Assert.That(text).IsEqualTo("beforeafter");
    }

    // Same refusal html gets, and for the same reason: the token replaces the block it sits in, so
    // text sharing that block would be discarded rather than kept.
    [Test]
    public async Task OpenXmlTokenPropertySharingItsBlockThrows()
    {
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () => await RenderText(
                "leading {{ Value }}",
                new TokenModel
                {
                    Value = new OpenXmlToken(_ => [new Paragraph(new Run(new Text("x")))])
                }));
        await Assert.That(exception!.Message).Contains("OpenXmlToken");
        await Assert.That(exception.Message).Contains("alone in that block");
    }

    // The one token left with no answer here. It mutates the paragraph a docx template already
    // had, and markdown has none to hand it until after the parse - by which point the paragraph
    // is Markdig's, not the template's.
    [Test]
    public async Task MutateTokenPropertyThrowsInMarkdownFlow()
    {
        var exception = await Assert.ThrowsAsync<ParchmentRenderException>(
            async () => await RenderText(
                "{{ Value }}",
                new TokenModel
                {
                    Value = new MutateToken((_, _) => { })
                }));
        await Assert.That(exception!.Message).Contains("MutateToken");
        await Assert.That(exception.Message).Contains("RegisterDocxTemplate");
    }

    static byte[] OnePixelPng() =>
        Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=");

    // The policies are OpenXmlHtml's and belong to the store, so they cover every template
    // registered on it rather than being decided per render.
    [Test]
    public async Task ImagePoliciesAreSetOnTheStore()
    {
        #region ImagePolicies

        var store = new TemplateStore
        {
            LocalImages = OpenXmlHtml.ImagePolicy.SafeDirectories("C:/assets/branding"),
            WebImages = OpenXmlHtml.ImagePolicy.Deny()
        };

        #endregion

        await Assert.That(store.LocalImages).IsNotSameReferenceAs(store.WebImages);
        await Assert.That(store.WebImages).IsNotNull();
    }
}
