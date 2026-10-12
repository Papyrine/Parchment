using DocumentFormat.OpenXml.Validation;

/// <summary>
/// A run carries one <c>w:rStyle</c>, and a link, emphasis, a classed span and inline code can each
/// ask for one. Which wins goes by what asked - emphasis, then the span, then the code, then the
/// link - and not by which of them encloses the other, so both nesting orders give the same run.
/// </summary>
public class CharacterStylePrecedenceTests
{
    record Decoration(string Name, string Open, string Close, int Rank = 0, string? Style = null, bool IsLink = false)
    {
        public string Wrap(string content) =>
            Open + content + Close;
    }

    const string url = "https://e.example";

    static readonly Decoration[] wrappers =
    [
        new("link", "[", $"]({url})", 1, "Hyperlink", true),
        new("styled link", "[", $"]({url}){{.LinkStyle}}", 1, "LinkStyle", true),
        new("bold", "**", "**"),
        new("styled italic", "_", "_{.EmphasisStyle}", 4, "EmphasisStyle"),
        new("html underline", "<u>", "</u>"),
        new("classed span", "<span class=\"SpanStyle\">", "</span>", 3, "SpanStyle"),
        new("classless span", "<span>", "</span>")
    ];

    // Nothing nests inside these, so each is only ever the innermost.
    static readonly Decoration[] leaves =
    [
        new("text", "", ""),
        new("code", "`", "`"),
        new("styled code", "`", "`{.CodeStyle}", 2, "CodeStyle")
    ];

    public static IEnumerable<(string first, string second, string leaf)> Pairs()
    {
        for (var i = 0; i < wrappers.Length; i++)
        {
            for (var j = i + 1; j < wrappers.Length; j++)
            {
                // A link cannot hold a link.
                if (wrappers[i].IsLink && wrappers[j].IsLink)
                {
                    continue;
                }

                foreach (var leaf in leaves)
                {
                    yield return (wrappers[i].Name, wrappers[j].Name, leaf.Name);
                }
            }
        }
    }

    [Test]
    [MethodDataSource(nameof(Pairs))]
    public async Task BothNestingOrdersGiveTheSameRun(string first, string second, string leaf)
    {
        var a = wrappers.Single(_ => _.Name == first);
        var b = wrappers.Single(_ => _.Name == second);
        var inner = leaves.Single(_ => _.Name == leaf);
        var content = inner.Wrap("x");

        var aOutside = await RunProperties(a.Wrap(b.Wrap(content)));
        var bOutside = await RunProperties(b.Wrap(a.Wrap(content)));

        await Assert.That(bOutside.OuterXml).IsEqualTo(aOutside.OuterXml);

        Decoration[] present = [a, b, inner];
        var expected = present
            .Where(_ => _.Style != null)
            .OrderByDescending(_ => _.Rank)
            .FirstOrDefault()
            ?.Style;
        await Assert.That(aOutside.RunStyle?.Val?.Value).IsEqualTo(expected);
    }

    public static IEnumerable<(string outer, string inner)> AroundAnAutolink()
    {
        var around = wrappers.Where(_ => !_.IsLink).ToList();
        foreach (var outer in around)
        {
            yield return (outer.Name, "");
            foreach (var inner in around)
            {
                if (inner != outer)
                {
                    yield return (outer.Name, inner.Name);
                }
            }
        }
    }

    // An autolink is added as a finished Hyperlink rather than as runs a link then wraps, so it is
    // held to the same link written out in full.
    [Test]
    [MethodDataSource(nameof(AroundAnAutolink))]
    public async Task AutolinkMatchesTheSameLinkWrittenOut(string outer, string inner)
    {
        var autolink = $"<{url}>";
        var link = $"[{url}]({url})";
        foreach (var name in new[] { inner, outer })
        {
            var decoration = wrappers.SingleOrDefault(_ => _.Name == name);
            if (decoration == null)
            {
                continue;
            }

            autolink = decoration.Wrap(autolink);
            link = decoration.Wrap(link);
        }

        var fromAutolink = await RunProperties(autolink, url);
        var fromLink = await RunProperties(link, url);

        await Assert.That(fromAutolink.OuterXml).IsEqualTo(fromLink.OuterXml);
    }

    [Test]
    [Arguments("<b><https://e.example></b>")]
    [Arguments("<strong><user@e.example></strong>")]
    [Arguments("<b>see https://e.example</b>")]
    [Arguments("<b>[text](https://e.example)</b>")]
    [Arguments("**<https://e.example>**")]
    public async Task BoldReachesTheLink(string markdown)
    {
        var paragraph = RendererHarness.RenderMarkdown($"before {markdown} after\n")
            .OfType<Paragraph>()
            .Single();
        var run = paragraph.Descendants<Hyperlink>().Single().Elements<Run>().Single();

        await Assert.That(run.RunProperties!.Bold).IsNotNull();
        await Assert.That(run.RunProperties.RunStyle!.Val?.Value).IsEqualTo("Hyperlink");
        await Assert.That(paragraph.Elements<Run>().Last().RunProperties?.Bold).IsNull();
    }

    [Test]
    [Arguments("<i>", "</i>", "i")]
    [Arguments("<u>", "</u>", "u")]
    [Arguments("<s>", "</s>", "strike")]
    [Arguments("<sup>", "</sup>", "vertAlign")]
    public async Task InlineHtmlReachesAnAutolink(string open, string close, string property)
    {
        var properties = await RunProperties($"{open}<{url}>{close}", url);

        var children = properties.ChildElements.Select(_ => _.LocalName);
        await Assert.That(string.Join(' ', children)).IsEqualTo($"rStyle {property}");
    }

    [Test]
    public async Task ClassedSpanStylesAnAutolink()
    {
        var properties = await RunProperties($"<span class=\"SpanStyle\"><{url}></span>", url);

        await Assert.That(properties.RunStyle!.Val?.Value).IsEqualTo("SpanStyle");
    }

    static async Task<RunProperties> RunProperties(string markdown, string text = "x")
    {
        var paragraph = RendererHarness.RenderMarkdown($"before {markdown} after\n")
            .OfType<Paragraph>()
            .Single();

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator
            .Validate(paragraph)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();

        var run = paragraph
            .Descendants<Run>()
            .Single(_ => _.InnerText == text);
        return run.RunProperties ?? new();
    }
}
