using DocumentFormat.OpenXml.Validation;

/// <summary>
/// Formatting reaches a run from whichever inline closes first, so the order the properties arrive
/// in follows the nesting. <c>w:rPr</c> is a strict sequence and holds each of these once, whatever
/// that order was.
/// </summary>
public class NestedFormattingTests
{
    [Test]
    // Emphasis inside emphasis
    [Arguments("**_x_**", "b i")]
    [Arguments("_**x**_", "b i")]
    [Arguments("***x***", "b i")]
    [Arguments("___x___", "b i")]
    [Arguments("*_x_*", "i")]
    [Arguments("**__x__**", "b")]
    [Arguments("~~**x**~~", "b strike")]
    [Arguments("**~~x~~**", "b strike")]
    [Arguments("==**x**==", "b highlight")]
    [Arguments("**==x==**", "b highlight")]
    [Arguments("++*x*++", "i u")]
    [Arguments("*++x++*", "i u")]
    [Arguments("==++x++==", "highlight u")]
    [Arguments("++==x==++", "highlight u")]
    [Arguments("~~==x==~~", "strike highlight")]
    [Arguments("==~~x~~==", "strike highlight")]
    [Arguments("^**x**^", "b vertAlign")]
    [Arguments("**^x^**", "b vertAlign")]
    [Arguments("~++x++~", "u vertAlign")]
    [Arguments("++~x~++", "u vertAlign")]
    [Arguments("^~x~^", "vertAlign")]
    [Arguments("++==~~***x***~~==++", "b i strike highlight u")]
    [Arguments("~~==++***x***++==~~", "b i strike highlight u")]
    // Inline html inside inline html
    [Arguments("<u><b>x</b></u>", "b u")]
    [Arguments("<b><u>x</u></b>", "b u")]
    [Arguments("<i><b>x</b></i>", "b i")]
    [Arguments("<s><em>x</em></s>", "i strike")]
    [Arguments("<sup><del>x</del></sup>", "strike vertAlign")]
    [Arguments("<sub><u>x</u></sub>", "u vertAlign")]
    [Arguments("<sup><u><s><i><b>x</b></i></s></u></sup>", "b i strike u vertAlign")]
    [Arguments("<b><strong>x</strong></b>", "b")]
    [Arguments("<s><del>x</del></s>", "strike")]
    [Arguments("<sub><sup>x</sup></sub>", "vertAlign")]
    // The two together
    [Arguments("**<u>x</u>**", "b u")]
    [Arguments("<u>**x**</u>", "b u")]
    [Arguments("<sup>**x**</sup>", "b vertAlign")]
    [Arguments("**<sup>x</sup>**", "b vertAlign")]
    [Arguments("<s>*x*</s>", "i strike")]
    [Arguments("==<b>x</b>==", "b highlight")]
    [Arguments("<u>==x==</u>", "highlight u")]
    [Arguments("<b>**x**</b>", "b")]
    [Arguments("**<b>x</b>**", "b")]
    [Arguments("<i>*x*</i>", "i")]
    [Arguments("<u>++x++</u>", "u")]
    [Arguments("++<u>x</u>++", "u")]
    [Arguments("<s>~~x~~</s>", "strike")]
    [Arguments("<sup>~x~</sup>", "vertAlign")]
    [Arguments("^<sub>x</sub>^", "vertAlign")]
    // Code spans bring rFonts, which sits ahead of every toggle
    [Arguments("**`x`**", "rFonts b")]
    [Arguments("==_`x`_==", "rFonts i highlight")]
    [Arguments("<u>`x`</u>", "rFonts u")]
    [Arguments("<u>**`x`**</u>", "rFonts b u")]
    // Character styles, which lead the sequence
    [Arguments("<u><span class=\"Lead\">x</span></u>", "rStyle u")]
    [Arguments("<span class=\"Lead\"><b>x</b></span>", "rStyle b")]
    [Arguments("**<span class=\"Lead\">x</span>**", "rStyle b")]
    [Arguments("<span class=\"Lead\">==*x*==</span>", "rStyle i highlight")]
    [Arguments("++**x**{.Lead}++", "rStyle b u")]
    [Arguments("<u>**x**{.Lead}</u>", "rStyle b u")]
    [Arguments("<u>`x`{.Lead}</u>", "rStyle rFonts u")]
    // Link text sits in a Hyperlink and already holds the link's rStyle
    [Arguments("<u>[x](https://example.com)</u>", "rStyle u")]
    [Arguments("<u>**[x](https://example.com)**</u>", "rStyle b u")]
    [Arguments("**<u>[x](https://example.com)</u>**", "rStyle b u")]
    [Arguments("[<sup>*x*</sup>](https://example.com)", "rStyle i vertAlign")]
    public async Task RunPropertiesKeepSchemaOrder(string markdown, string expected)
    {
        var paragraph = RendererHarness.RenderMarkdown($"before {markdown} after\n")
            .OfType<Paragraph>()
            .Single();
        var run = paragraph
            .Descendants<Run>()
            .Single(_ => _.InnerText == "x");

        var children = run.RunProperties!.ChildElements.Select(_ => _.LocalName);
        await Assert.That(string.Join(' ', children)).IsEqualTo(expected);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2013);
        var errors = validator
            .Validate(paragraph)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    [Arguments("<sub><sup>x</sup></sub>", "superscript")]
    [Arguments("<sup><sub>x</sub></sup>", "subscript")]
    [Arguments("^<sub>x</sub>^", "subscript")]
    [Arguments("~<sup>x</sup>~", "superscript")]
    public async Task InnermostVerticalAlignmentWins(string markdown, string expected)
    {
        var paragraph = RendererHarness.RenderMarkdown($"before {markdown} after\n")
            .OfType<Paragraph>()
            .Single();
        var run = paragraph
            .Descendants<Run>()
            .Single(_ => _.InnerText == "x");

        var alignment = run.RunProperties!.Elements<VerticalTextAlignment>().Single();
        await Assert.That(alignment.Val!.InnerText).IsEqualTo(expected);
    }
}
