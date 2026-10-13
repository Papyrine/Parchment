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

    [ParchmentBindable]
    public partial class ShadowedAnnotatedModel
    {
        public required AnnotatedPerson Customer { get; init; }
        public required List<PlainPerson> Customers { get; init; }
        public List<PlainPerson> Nobody { get; init; } = [];
    }

    public class AnnotatedPerson
    {
        [Html]
        public required string Notes { get; init; }

        public string Name { get; init; } = "Root";
    }

    public class PlainPerson
    {
        public required string Notes { get; init; }
    }

    static async Task<Body> RenderShadowed(string template)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ShadowedAnnotatedModel>(template, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ShadowedAnnotatedModel
            {
                Customer = new()
                {
                    Notes = "<b>ROOT</b>"
                },
                Customers =
                [
                    new()
                    {
                        Notes = "<b>one</b>"
                    }
                ]
            },
            stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        return (Body) document.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    // `customer.Notes` spells the root's annotated `Customer.Notes`, but inside the loop it is the
    // item's member, which carries no annotation: it is text, and is escaped as text.
    [Test]
    public async Task LoopVariableNamedAfterAnnotatedRootMemberStaysText()
    {
        var body = await RenderShadowed(
            """
            {{ Customer.Notes }}

            {% for customer in Customers %}
            {{ customer.Notes }}

            {% endfor %}
            """);

        var paragraphs = body.Elements<Paragraph>()
            .Where(_ => _.InnerText.Length > 0)
            .ToList();

        await Assert.That(paragraphs.Select(_ => _.InnerText)).IsEquivalentTo(["ROOT", "<b>one</b>"]);
        await Assert.That(paragraphs[0].Descendants<Bold>().Any()).IsTrue();
        await Assert.That(paragraphs[1].Descendants<Bold>().Any()).IsFalse();
    }

    // Fluid finds a loop variable by its exact spelling and a model member without regard to
    // case, so the root's own spelling still reaches the root, and its annotation, inside the loop.
    [Test]
    public async Task LoopVariableShadowsOnlyItsOwnSpelling()
    {
        var body = await RenderShadowed(
            """
            {% for customer in Customers %}
            {{ Customer.Notes }}

            {% endfor %}
            """);

        var paragraph = body.Elements<Paragraph>().Single(_ => _.InnerText.Length > 0);

        await Assert.That(paragraph.InnerText).IsEqualTo("ROOT");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsTrue();
    }

    static List<Paragraph> Filled(Body body) =>
        body.Elements<Paragraph>()
            .Where(_ => _.InnerText.Length > 0)
            .ToList();

    // A raw region is printed as written. The rewrite used to reach into it and add the filter, so
    // the document showed `{{ Customer.Notes | html }}`, which nobody wrote.
    [Test]
    [Arguments("{% raw %}{{ Customer.Notes }}{% endraw %}")]
    [Arguments("{%- raw -%}{{ Customer.Notes }}{%- endraw -%}")]
    public async Task TokenInsideRawIsPrintedAsWritten(string template)
    {
        var body = await RenderShadowed(template);

        await Assert.That(Filled(body).Select(_ => _.InnerText)).IsEquivalentTo(["{{ Customer.Notes }}"]);
    }

    // What reads as a loop tag inside a raw or comment region opens no loop, so the token after it
    // is the root's member and carries its annotation.
    [Test]
    [Arguments("{% raw %}{% for Customer in Customers %}{% endraw %}", 2)]
    [Arguments("{% comment %}{% for Customer in Customers %}{% endcomment %}", 1)]
    [Arguments("{%- comment -%}{% for Customer in Customers %}{%- endcomment -%}", 1)]
    [Arguments("{% # {% for Customer in Customers %}", 1)]
    public async Task LoopTagInsideRawOrCommentBindsNothing(string region, int expected)
    {
        var body = await RenderShadowed(region + "\n\n{{ Customer.Notes }}");

        var paragraphs = Filled(body);
        await Assert.That(paragraphs.Count).IsEqualTo(expected);
        await Assert.That(paragraphs[^1].InnerText).IsEqualTo("ROOT");
        await Assert.That(paragraphs[^1].Descendants<Bold>().Any()).IsTrue();
    }

    // Nor does an `{% endfor %}` in such a region close one: the token after it is still the item.
    [Test]
    [Arguments("{% raw %}{% endfor %}{% endraw %}", 2)]
    [Arguments("{% comment %}{% endfor %}{% endcomment %}", 1)]
    public async Task EndTagInsideRawOrCommentClosesNothing(string region, int expected)
    {
        var body = await RenderShadowed(
            "{% for Customer in Customers %}" + region + "\n\n{{ Customer.Notes }}\n\n{% endfor %}");

        var paragraphs = Filled(body);
        await Assert.That(paragraphs.Count).IsEqualTo(expected);
        await Assert.That(paragraphs[^1].InnerText).IsEqualTo("<b>one</b>");
        await Assert.That(paragraphs[^1].Descendants<Bold>().Any()).IsFalse();
    }

    // The else of a loop renders when there was nothing to iterate, so the loop variable is not
    // bound in it and the same spelling reads the root member.
    [Test]
    public async Task LoopVariableIsNotBoundInTheElseOfItsLoop()
    {
        var body = await RenderShadowed(
            "{% for Customer in Nobody %}{{ Customer.Notes }}{% else %}{{ Customer.Notes }}{% endfor %}");

        var paragraph = Filled(body).Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("ROOT");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsTrue();
    }

    // An else inside the loop that belongs to an if is not the loop's, and ends nothing.
    [Test]
    public async Task ElseOfANestedIfDoesNotEndTheLoop()
    {
        var body = await RenderShadowed(
            "{% for Customer in Customers %}{% if false %}x{% else %}{{ Customer.Notes }}{% endif %}{% endfor %}");

        var paragraph = Filled(body).Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("<b>one</b>");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsFalse();
    }

    // The validator agrees: in the else, `Customer.Name` is the root's member, which the item type
    // does not have. Bound to the item it failed registration for a template that renders.
    [Test]
    public async Task ElseOfALoopIsValidatedAgainstTheEnclosingScope()
    {
        var body = await RenderShadowed(
            "{% for Customer in Nobody %}x{% else %}{{ Customer.Name }}{% endfor %}");

        await Assert.That(Filled(body).Single().InnerText).IsEqualTo("Root");
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

    // Every punctuation character in a value is escaped, and Markdig ends a literal at each escape,
    // so a value arrives as one run per fragment. Fragments formatted alike fold back into one run,
    // and a fragment is never folded into a neighbour formatted any differently.
    [Test]
    public async Task PunctuatedValueInsideEmphasisFoldsToOneRunPerFormat()
    {
        var body = await Render(
            """
            On-call: **{{ Title }}** then _{{ Title }}_**_{{ Details }}_**, <u>{{ Title }}</u>++{{ Details }}++ and ==**{{ Title }}**==**{{ Details }}** (end).

            | Name | Note |
            | --- | --- |
            | **{{ Title }}** | ~~{{ Details }}~~ {{ Title }} |
            """,
            new()
            {
                Title = "Item-1 (rev. 1.2)",
                Details = "a_b*c [d] #e!"
            });

        var runs = body.Descendants<Run>()
            .Select(_ => $"{string.Join(' ', _.RunProperties?.Select(property => property.LocalName) ?? [])}|{_.InnerText}")
            .ToList();

        await Assert.That(runs)
            .IsEquivalentTo(
            [
                "|On-call: ",
                "b|Item-1 (rev. 1.2)",
                "| then ",
                "i|Item-1 (rev. 1.2)",
                "b i|a_b*c [d] #e!",
                "|, ",
                "u|Item-1 (rev. 1.2)a_b*c [d] #e!",
                "| and ",
                "b highlight|Item-1 (rev. 1.2)",
                "b|a_b*c [d] #e!",
                "| (end).",
                "b|Name",
                "b|Note",
                "b|Item-1 (rev. 1.2)",
                "strike|a_b*c [d] #e!",
                "| Item-1 (rev. 1.2)"
            ]);
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

    const string entitiesAndDelimiters = "AT&T &amp; &copy; &#169; <Research> a < b [open ![image <<x>> \"Labs\" it's";

    // The renderer turns an entity in the template into its character and prints a bracket that
    // never closed. Neither is to reach into a value: one holding "&amp;" holds five characters,
    // and they print as five - once, not decoded, and not decoded twice.
    [Test]
    [Arguments("{0}", "{0}")]
    [Arguments("**{0}**", "{0}")]
    [Arguments("# {0}", "{0}")]
    [Arguments("- {0}", "{0}")]
    [Arguments("| Head |\n| --- |\n| {0} |", "Head{0}")]
    [Arguments("[{0}](https://example.com)", "{0}")]
    [Arguments("<u>{0}</u>", "{0}")]
    public async Task EntitiesAndDelimitersInAValuePrintAsThemselves(string container, string expected)
    {
        // Quotes are left to smarty pants by design, so they are kept out of this one.
        var value = entitiesAndDelimiters.Replace("\"Labs\" it's", "end");
        var body = await Render(
            string.Format(container, "{{ Title }}"),
            new()
            {
                Title = value
            });

        await Assert.That(body.InnerText).IsEqualTo(string.Format(expected, value));
    }

    // escape_xml opts out of the encoder and writes entities in its place, so in running text it is
    // the renderer that turns them back - a value through it used to lose every character the
    // filter had touched. What the filter leaves alone is markdown source, which is where the open
    // bracket comes from; it prints too.
    [Test]
    [Arguments("{0}", "{0}")]
    [Arguments("Company: {0}.", "Company: {0}.")]
    [Arguments("**{0}**", "{0}")]
    [Arguments("# {0}", "{0}")]
    [Arguments("- {0}", "{0}")]
    [Arguments("| Head |\n| --- |\n| {0} |", "Head{0}")]
    [Arguments("<u>{0}</u>", "{0}")]
    public async Task EscapeXmlValuePrintsAsItself(string container, string expected)
    {
        // Angle quotes need two angle brackets in the source, and the filter has made entities of
        // them, so those print as typed as well.
        var body = await Render(
            string.Format(container, "{{ Title | escape_xml }}"),
            new()
            {
                Title = entitiesAndDelimiters
            });

        await Assert.That(body.InnerText).IsEqualTo(string.Format(expected, entitiesAndDelimiters));
    }

    // The docx flow puts the value in a run and never parses it, so there is nothing to decode.
    [Test]
    public async Task DocxSubstitutionKeepsEntitiesAndDelimiters()
    {
        var body = await RenderDocx(
            "{{ Title }}",
            new()
            {
                Title = entitiesAndDelimiters
            });

        await Assert.That(body.Elements<Paragraph>().First().InnerText).IsEqualTo(entitiesAndDelimiters);
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

    // escape_xml opts out of the encoder, so the line break promise is its own to keep. A newline
    // left in the source is whitespace at best, and a blank line ends the html block the value was
    // escaped for — the rest of a table then spills out of it as markdown.
    [Test]
    public async Task EscapeXmlBlankLineStaysInsideItsHtmlBlock()
    {
        var body = await Render(
            """
            <table>
            <tr><td>{{ Details | escape_xml }}</td><td>after</td></tr>
            </table>
            """,
            new()
            {
                Details = "one\r\n\r\ntwo & <three>"
            });

        var cells = body.Descendants<TableCell>().ToList();
        await Assert.That(cells.Count).IsEqualTo(2);
        await Assert.That(cells[0].InnerText).IsEqualTo("onetwo & <three>");
        await Assert.That(cells[0].Descendants<Break>().Count()).IsEqualTo(2);
        await Assert.That(cells[1].InnerText).IsEqualTo("after");
    }

    [Test]
    [Arguments("one\ntwo")]
    [Arguments("one\r\ntwo")]
    [Arguments("one\rtwo")]
    public async Task EscapeXmlNewlineBecomesABreak(string value)
    {
        var body = await Render(
            "{{ Title | escape_xml }}",
            new()
            {
                Title = value
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(paragraph.InnerText).IsEqualTo("onetwo");
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

    // Here `html` parks its value and writes a placeholder, so a filter after it was handed the
    // placeholder rather than the value: `{{ Details | html | upcase }}` printed PARCHMENT-TOKEN-0,
    // and `append` left the placeholder sharing its block, which failed the render with a message
    // about a block the template never wrote.
    [Test]
    [Arguments("{{ Details | html | upcase }}", "upcase", "Details | html | upcase")]
    [Arguments("{{ Details | html | append: ' tail' }}", "append", "Details | html | append")]
    [Arguments("{{ Details | upcase | html | downcase | strip }}", "downcase", "Details | upcase | html | downcase | strip")]
    [Arguments("{% assign shout = Details | html | upcase %}{{ shout }}", "upcase", "Details | html | upcase")]
    [Arguments("{% capture held %}{{ Details | html | upcase }}{% endcapture %}{{ held }}", "upcase", "Details | html | upcase")]
    [Arguments("{% for item in Items %}{{ item | html | upcase }}{% endfor %}", "upcase", "item | html | upcase")]
    [Arguments("{% if Title %}x{% else %}{{ Details | html | upcase }}{% endif %}", "upcase", "Details | html | upcase")]
    public async Task FilterAfterHtmlFilterFailsRegistration(string template, string following, string chain)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();

        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => store.RegisterMarkdownTemplate<EscapeModel>(template, styleSource));

        await Assert.That(exception.Message).Contains($"Filter '{following}' is applied after 'html'");
        await Assert.That(exception.Message).Contains($"Move '{following}' before 'html'");
        await Assert.That(exception.TokenSource).IsEqualTo(chain);
    }

    [Test]
    public async Task FilterBeforeHtmlFilterIsApplied()
    {
        var body = await Render(
            "{{ Details | upcase | html }}",
            new()
            {
                Details = "<b>bold</b> text"
            });

        await Assert.That(body.InnerText).IsEqualTo("BOLD TEXT");
        await Assert.That(body.Descendants<Run>().Single(_ => _.RunProperties?.Bold != null).InnerText).IsEqualTo("BOLD");
    }

    // The other three are not token filters in this flow: each writes markdown source, and source is
    // text a later filter can still work on. The docx flow refuses these same chains, because there
    // the value has already become Word content.
    [Test]
    public async Task FilterAfterMarkdownFilterTransformsItsSource()
    {
        var body = await Render(
            "{{ Details | markdown | upcase }}",
            new()
            {
                Details = "**bold** text"
            });

        await Assert.That(body.InnerText).IsEqualTo("BOLD TEXT");
        await Assert.That(body.Descendants<Run>().Single(_ => _.RunProperties?.Bold != null).InnerText).IsEqualTo("BOLD");
    }

    [Test]
    [Arguments("bullet_list")]
    [Arguments("numbered_list")]
    public async Task FilterAfterListFilterTransformsItsSource(string listFilter)
    {
        var body = await Render(
            $"{{{{ Items | {listFilter} | upcase }}}}",
            new()
            {
                Items = ["one", "two"]
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(string.Join(',', paragraphs.Select(_ => _.InnerText))).IsEqualTo("ONE,TWO");
        await Assert.That(paragraphs.All(_ => _.ParagraphProperties?.NumberingProperties != null)).IsTrue();
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

    // Whitespace control and a line break inside the braces are spellings of the same token, so
    // the annotation reaches them too. They used to print the member's source as text.
    [Test]
    [Arguments("{{- Marked -}}")]
    [Arguments("{{- Marked }}")]
    [Arguments("{{ Marked -}}")]
    [Arguments("{{-Marked-}}")]
    [Arguments("{{\n  Marked\n}}")]
    [Arguments("{{-\r\n  Marked\r\n-}}")]
    public async Task AnnotationIsHonouredInEverySpelling(string template)
    {
        var body = await RenderAnnotated(
            template,
            new()
            {
                Marked = "## Heading"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("Heading");
        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("Heading2");
    }

    // `Customer['Notes']` is `Customer.Notes` to Fluid, so it is the same annotated member.
    [Test]
    [Arguments("{{ Customer['Notes'] }}")]
    [Arguments("{{ Customer[\"Notes\"] }}")]
    [Arguments("{{- Customer['Notes'] -}}")]
    [Arguments("{{\n  Customer['Notes']\n}}")]
    public async Task AnnotationIsHonouredThroughBracketAccess(string template)
    {
        var body = await RenderShadowed(template);

        var paragraph = body.Elements<Paragraph>().Single(_ => _.InnerText.Length > 0);
        await Assert.That(paragraph.InnerText).IsEqualTo("ROOT");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsTrue();
    }

    // The rewritten token keeps its dashes: the lines around it still join it.
    [Test]
    public async Task AnnotatedTokenKeepsItsWhitespaceControl()
    {
        var body = await RenderAnnotated(
            """
            Before:
            {{- Marked -}}
            :after
            """,
            new()
            {
                Marked = "*slanted*"
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await Assert.That(paragraph.InnerText).IsEqualTo("Before:slanted:after");
        await Assert.That(paragraph.Descendants<Italic>().Any()).IsTrue();
    }

    // The loop variable still wins in the other spellings: the item's member carries no annotation.
    [Test]
    [Arguments("{{- customer.Notes -}}")]
    [Arguments("{{ customer['Notes'] }}")]
    [Arguments("{{\n  customer.Notes\n}}")]
    public async Task LoopVariableStaysTextInEverySpelling(string token)
    {
        var body = await RenderShadowed($"{{% for customer in Customers %}}\n\n{token}\n\n{{% endfor %}}");

        var paragraph = body.Elements<Paragraph>().Single(_ => _.InnerText.Length > 0);
        await Assert.That(paragraph.InnerText).IsEqualTo("<b>one</b>");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsFalse();
    }

    // A filter still makes the value the filter's output, however the token is spelled.
    [Test]
    [Arguments("{{- Customer.Notes | upcase -}}")]
    [Arguments("{{ Customer['Notes'] | upcase }}")]
    [Arguments("{{\n  Customer.Notes\n  | upcase\n}}")]
    public async Task FilteredAnnotatedMemberIsStillEscapedInEverySpelling(string template)
    {
        var body = await RenderShadowed(template);

        var paragraph = body.Elements<Paragraph>().Single(_ => _.InnerText.Length > 0);
        await Assert.That(paragraph.InnerText).IsEqualTo("<B>ROOT</B>");
        await Assert.That(paragraph.Descendants<Bold>().Any()).IsFalse();
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

    // escape_xml writes its break as html only where the value becomes source. Here it lands in a
    // run, where the tag would print, so the newline is left for LineBreaks like any other.
    [Test]
    public async Task DocxEscapeXmlNewlineBecomesABreak()
    {
        var body = await RenderDocx(
            "{{ Title | escape_xml }}",
            new()
            {
                Title = "one\ntwo"
            });

        var paragraph = body.Elements<Paragraph>().First();
        await Assert.That(paragraph.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(paragraph.InnerText).IsEqualTo("onetwo");
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

    // Markdown has no empty emphasis: the delimiters would print, and a line holding only `****`
    // or `___` is a thematic break. A blank value takes the emphasis written around it with it.
    [Test]
    [Arguments("**{{ Title }}**", "")]
    [Arguments("__{{ Title }}__", "")]
    [Arguments("*{{ Title }}*", "")]
    [Arguments("_{{ Title }}_", "")]
    [Arguments("~~{{ Title }}~~", "")]
    [Arguments("`{{ Title }}`", "")]
    [Arguments("***{{ Title }}***", "")]
    [Arguments("~~**{{ Title }}**~~", "")]
    [Arguments("**{{ Title }}{{ Details }}**", "")]
    [Arguments("**{{ Title }} {{ Details }}**", "")]
    [Arguments("**{{ Title }}**", " ")]
    [Arguments("**{{ Title }}**", " \t ")]
    [Arguments("_{{ Title | upcase }}_", "")]
    [Arguments("**{{ Title | markdown }}**", "")]
    public async Task EmphasisAroundABlankValuePrintsNothing(string emphasis, string value)
    {
        var body = await Render(
            $"before {emphasis} after\n\n{emphasis}\n\nlast",
            new()
            {
                Title = value,
                Details = value
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs.Select(_ => string.Join(' ', _.InnerText.Split(' ', StringSplitOptions.RemoveEmptyEntries))))
            .IsEquivalentTo(["before after", "last"]);
        await Assert.That(body.Descendants<ParagraphBorders>().Any()).IsFalse();
    }

    [Test]
    public async Task EmphasisAroundABlankValueInALoop()
    {
        var body = await Render(
            "{% assign nothing = nil %}{% for item in Items %}**{{ item }}** __{{ nothing }}__\n\n{% endfor %}last",
            new()
            {
                Items = ["", "kept", " "]
            });

        await Assert.That(body.InnerText.Replace(" ", "")).IsEqualTo("keptlast");
        await Assert.That(body.Descendants<Bold>().Single().Ancestors<Run>().Single().InnerText).IsEqualTo("kept");
        await Assert.That(body.Descendants<ParagraphBorders>().Any()).IsFalse();
    }

    [Test]
    public async Task EmphasisAroundABlankValueInsideOtherBlocks()
    {
        var body = await Render(
            """
            # **{{ Title }}**

            - **{{ Title }}** item
            - _{{ Title }}_

            | A | B |
            | --- | --- |
            | **{{ Title }}** | ~~{{ Title }}~~ after |
            """,
            new());

        await Assert.That(body.InnerText).IsEqualTo("itemABafter");
        await Assert.That(body.Descendants<Strike>().Any()).IsFalse();
        await Assert.That(body.Descendants<Italic>().Any()).IsFalse();
        await Assert.That(body.Descendants<ParagraphBorders>().Any()).IsFalse();
    }

    [Test]
    public async Task EmphasisAroundAValueIsKept()
    {
        var body = await Render(
            "**{{ Title }}** and _{{ Details }}_",
            new()
            {
                Title = "bold",
                Details = "italic"
            });

        await Assert.That(body.InnerText).IsEqualTo("bold and italic");
        await Assert.That(body.Descendants<Bold>().Single().Ancestors<Run>().Single().InnerText).IsEqualTo("bold");
        await Assert.That(body.Descendants<Italic>().Single().Ancestors<Run>().Single().InnerText).IsEqualTo("italic");
    }

    // The runs either side of the token here close one emphasis and open another. Neither is
    // around the token, so both stay.
    [Test]
    [Arguments("**a**{{ Title }}*b*")]
    [Arguments("**a** {{ Title }} *b*")]
    public async Task NeighbouringEmphasisSurvivesABlankValue(string template)
    {
        var body = await Render(template, new());

        await Assert.That(body.InnerText.Replace(" ", "")).IsEqualTo("ab");
        await Assert.That(body.Descendants<Bold>().First().Ancestors<Run>().Single().InnerText).IsEqualTo("a");
    }

    // Only delimiters around a token are dropped. Ones the author typed with nothing between them,
    // in the template or in a value that is markdown source, print as they always did.
    [Test]
    public async Task EmptyEmphasisWithoutATokenIsLeftAlone()
    {
        var body = await Render(
            "a **** b {{ Title }}\n\n{{ Details | markdown }}",
            new()
            {
                Details = "c **** d"
            });

        await Assert.That(body.Elements<Paragraph>().Select(_ => _.InnerText.Trim())).IsEquivalentTo(["a **** b", "c **** d"]);
    }

    [Test]
    public async Task EscapedDelimitersAroundABlankValueStillPrint()
    {
        var body = await Render(@"a \*{{ Title }}\* b", new());

        await Assert.That(body.InnerText).IsEqualTo("a ** b");
    }

    [Test]
    public async Task DocxMarkdownValueKeepsItsEmptyEmphasis()
    {
        var body = await RenderDocx(
            "{{ Details | markdown }}",
            new()
            {
                Details = "a **** b"
            });

        await Assert.That(body.InnerText).IsEqualTo("a **** b");
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

    // Every construct MarkdigPipeline reads at the start of a line. Whatever the value opens with,
    // it is one plain paragraph of that text.
    [Test]
    [Arguments("    indented code")]
    [Arguments("\tindented code")]
    [Arguments("        deeper")]
    [Arguments(" one space")]
    [Arguments("   three spaces")]
    [Arguments("    # indented heading")]
    [Arguments("```")]
    [Arguments("```cs")]
    [Arguments("~~~")]
    [Arguments("# heading")]
    [Arguments("###### heading")]
    [Arguments("#")]
    [Arguments("> quote")]
    [Arguments(">quote")]
    [Arguments("- item")]
    [Arguments("+ item")]
    [Arguments("* item")]
    [Arguments("-")]
    [Arguments("1. item")]
    [Arguments("1) item")]
    [Arguments("a. item")]
    [Arguments("A) item")]
    [Arguments("i. item")]
    [Arguments("IV) item")]
    [Arguments("---")]
    [Arguments("***")]
    [Arguments("___")]
    [Arguments("- - -")]
    [Arguments("* * *")]
    [Arguments("===")]
    [Arguments("| a | b |")]
    [Arguments("a | b")]
    [Arguments("+---+---+")]
    [Arguments("<div>block</div>")]
    [Arguments("<!-- comment -->")]
    [Arguments("<script>")]
    [Arguments("<?php")]
    [Arguments("<![CDATA[x]]>")]
    [Arguments("</div>")]
    [Arguments("[TOC]")]
    [Arguments("[SECTION]")]
    [Arguments("[TOC]{levels=1}")]
    [Arguments("[SECTION]{orientation=landscape}")]
    [Arguments("[label]: https://example.com")]
    [Arguments("[^1]: footnote")]
    [Arguments("*[HTML]: abbreviation")]
    [Arguments(": definition")]
    [Arguments("{.Heading1}")]
    [Arguments("$$")]
    [Arguments(":::")]
    public async Task BlockSyntaxOpeningAValuePrintsAsItself(string value)
    {
        foreach (var template in (string[])
                 [
                     "{{ Details }}",
                     "{{ Details }}\n",
                     "{{ Title }}{{ Details }}",
                     "before\n\n{{ Details }}\n\nafter"
                 ])
        {
            var body = await Render(
                template,
                new()
                {
                    Details = value
                });

            var paragraph = body.Elements<Paragraph>().Single(_ => _.InnerText != "before" && _.InnerText != "after");
            await AssertPlain(body, paragraph);
            await Assert.That(paragraph.InnerText).IsEqualTo(value.Replace("\t", ""));
            await Assert.That(paragraph.Descendants<TabChar>().Count()).IsEqualTo(value.Count(_ => _ == '\t'));
        }
    }

    // A later line of the value never starts a line of the source: its newline is a <br />, so
    // what follows is the middle of a paragraph whatever it reads as.
    [Test]
    [Arguments("    indented code")]
    [Arguments("\tindented code")]
    [Arguments("```")]
    [Arguments("# heading")]
    [Arguments("> quote")]
    [Arguments("- item")]
    [Arguments("1. item")]
    [Arguments("---")]
    [Arguments("===")]
    [Arguments("| a | b |")]
    [Arguments("|---|---|")]
    [Arguments("<div>")]
    [Arguments("[TOC]")]
    [Arguments("[SECTION]")]
    public async Task BlockSyntaxOnALaterLineOfAValuePrintsAsItself(string line)
    {
        foreach (var value in (string[])
                 [
                     $"text\n{line}",
                     $"text\r\n{line}",
                     $"text\n\n{line}",
                     $"text\n{line}\nmore",
                     $"\n{line}",
                     $"    lead\n{line}"
                 ])
        {
            var body = await Render(
                "{{ Details }}",
                new()
                {
                    Details = value
                });

            var paragraph = body.Elements<Paragraph>().Single();
            await AssertPlain(body, paragraph);
            await Assert.That(paragraph.InnerText).IsEqualTo(value.Replace("\r", "").Replace("\n", "").Replace("\t", ""));
        }
    }

    // The value is on the line after text, where = or - alone would underline that text into a
    // heading. Anything else is a continuation of the paragraph, which is where it belongs.
    [Test]
    [Arguments("===")]
    [Arguments("---")]
    [Arguments("=")]
    [Arguments("-")]
    [Arguments("    indented")]
    [Arguments("# heading")]
    [Arguments("> quote")]
    [Arguments("- item")]
    [Arguments("1. item")]
    [Arguments("```")]
    [Arguments("|---|---|")]
    [Arguments("<div>")]
    [Arguments("[TOC]")]
    public async Task BlockSyntaxOnTheLineAfterTextStaysInThatParagraph(string value)
    {
        var body = await Render(
            "Heading text\n{{ Details }}",
            new()
            {
                Details = value
            });

        var paragraph = body.Elements<Paragraph>().Single();
        await AssertPlain(body, paragraph);
        await Assert.That(paragraph.InnerText).StartsWith("Heading text");
        await Assert.That(paragraph.InnerText).EndsWith(value);
    }

    // A value that reads as a pipe-table row, over one that reads as its separator.
    [Test]
    public async Task TableSyntaxInAValueDoesNotMakeATable()
    {
        var body = await Render(
            "{{ Title }}\n{{ Details }}",
            new()
            {
                Title = "a | b",
                Details = "--- | ---"
            });

        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        await Assert.That(body.Elements<Paragraph>().Single().InnerText).Contains("a | b");
    }

    // The value the markers themselves are spelt with, in the places a marker is honoured: alone in
    // a paragraph, and with the attributes a template attaches to one.
    [Test]
    [Arguments("{{ Details }}", "[TOC]", "[TOC]")]
    [Arguments("{{ Details }}", " [TOC]", " [TOC]")]
    [Arguments("{{ Details }}{levels=1}", "[TOC]", "[TOC]")]
    [Arguments("{{ Details }}", "[SECTION]", "[SECTION]")]
    [Arguments("{{ Details }}{orientation=landscape}", "[SECTION]", "[SECTION]")]
    [Arguments("{{ Details }}{{ Title }}", "[TOC]", "[TOC]")]
    [Arguments("# Heading\n\n{{ Details }}\n\ntext", "[TOC]", "Heading[TOC]text")]
    public async Task MarkerInAValueIsNotAMarker(string template, string value, string expected)
    {
        var body = await Render(
            template,
            new()
            {
                Details = value
            });

        await Assert.That(body.Descendants<FieldChar>().Any()).IsFalse();
        await Assert.That(body.Descendants<FieldCode>().Any()).IsFalse();
        await Assert.That(body.Elements<Paragraph>().Any(_ => _.ParagraphProperties?.SectionProperties != null)).IsFalse();
        await Assert.That(body.InnerText).IsEqualTo(expected);
    }

    // The template's own markers are untouched by any of this.
    [Test]
    public async Task MarkerInTheTemplateIsStillAMarker()
    {
        var body = await Render(
            "[TOC]\n\n{{ Details }}\n\n[SECTION]{orientation=landscape}\n\nwide",
            new()
            {
                Details = "[TOC]"
            });

        await Assert.That(body.Descendants<FieldCode>().Count()).IsEqualTo(1);
        await Assert.That(body.Elements<Paragraph>().Count(_ => _.ParagraphProperties?.SectionProperties != null)).IsEqualTo(1);
        await Assert.That(body.Elements<Paragraph>().Count(_ => _.InnerText == "[TOC]")).IsEqualTo(1);
    }

    // The indent of a value is part of the value. It survives as the spaces it is, the way it does
    // in the docx flow, rather than being stripped as source indentation or read as a code block.
    [Test]
    [Arguments("{{ Details }}")]
    [Arguments("| a |\n| --- |\n| {{ Details }} |")]
    [Arguments("# {{ Details }}")]
    [Arguments("> {{ Details }}")]
    [Arguments("- {{ Details }}")]
    public async Task LeadingWhitespaceOfAValueSurvives(string template)
    {
        var body = await Render(
            template,
            new()
            {
                Details = "    four"
            });

        var texts = body.Descendants<Text>().Where(_ => _.Text.Contains("four")).ToList();
        await Assert.That(string.Concat(texts.Select(_ => _.Text))).IsEqualTo("    four");
        await Assert.That(texts[0].Space?.Value).IsEqualTo(SpaceProcessingModeValues.Preserve);
    }

    // The filter's marker is followed by the item, and five columns after a marker is an indented
    // code block inside the item.
    [Test]
    public async Task ListFilterItemKeepsItsIndentAndStaysAnItem()
    {
        var body = await Render(
            "{{ Items | bullet_list }}",
            new()
            {
                Items = ["     five", "\ttab", "[TOC]"]
            });

        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(3);
        await Assert.That(paragraphs[0].InnerText).IsEqualTo("     five");
        await Assert.That(paragraphs[2].InnerText).IsEqualTo("[TOC]");
        await Assert.That(paragraphs.All(_ => _.ParagraphProperties!.NumberingProperties != null)).IsTrue();
        await Assert.That(paragraphs.Select(_ => _.ParagraphProperties!.ParagraphStyleId?.Val?.Value).Distinct().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<FieldChar>().Any()).IsFalse();
    }

    [Test]
    [Arguments(" a", "&#32;a")]
    [Arguments("    a", "&#32;   a")]
    [Arguments("\ta", "&#9;a")]
    [Arguments(" \ta", "&#32;\ta")]
    [Arguments(" ", "&#32;")]
    [Arguments(" - a", @"&#32;\- a")]
    [Arguments("[TOC]", @"&#91;TOC\]")]
    [Arguments("a[b]", @"a\[b\]")]
    [Arguments("a b ", "a b ")]
    [Arguments("a\n    b", "a<br />    b")]
    public async Task EncoderWritesTheLeadingCharacterAsAReference(string value, string expected)
    {
        await Assert.That(MarkdownEncoder.Default.Encode(value)).IsEqualTo(expected);

        using var writer = new StringWriter();
        MarkdownEncoder.Default.Encode(writer, value);
        await Assert.That(writer.ToString()).IsEqualTo(expected);
    }

    // What the leading character of a value is carried as. Markdig parses a reference into an
    // inline of its own, which has to be rendered or the character is dropped from the document.
    [Test]
    public async Task CharacterReferenceInTheTemplatePrintsItsCharacter()
    {
        var body = await Render("Fish &amp; chips &#169;{{ Title }}", new());

        await Assert.That(body.InnerText).IsEqualTo("Fish & chips ©");
    }

    static async Task AssertPlain(Body body, Paragraph paragraph)
    {
        await Assert.That(body.Descendants<Table>().Any()).IsFalse();
        await Assert.That(body.Descendants<FieldChar>().Any()).IsFalse();
        await Assert.That(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value).IsNull();
        await Assert.That(paragraph.ParagraphProperties?.NumberingProperties).IsNull();
        await Assert.That(paragraph.ParagraphProperties?.SectionProperties).IsNull();
    }

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
