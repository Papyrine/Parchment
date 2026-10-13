public class SmartyPantInlineRendererTests
{
    [Test]
    [Arguments(SmartyPantType.Quote, "'")]
    [Arguments(SmartyPantType.LeftQuote, "\u2018")]
    [Arguments(SmartyPantType.RightQuote, "\u2019")]
    [Arguments(SmartyPantType.DoubleQuote, "\"")]
    [Arguments(SmartyPantType.LeftDoubleQuote, "\u201C")]
    [Arguments(SmartyPantType.RightDoubleQuote, "\u201D")]
    [Arguments(SmartyPantType.LeftAngleQuote, "\u00AB")]
    [Arguments(SmartyPantType.RightAngleQuote, "\u00BB")]
    [Arguments(SmartyPantType.Dash2, "\u2013")]
    [Arguments(SmartyPantType.Dash3, "\u2014")]
    [Arguments(SmartyPantType.Ellipsis, "\u2026")]
    public async Task EmitsExpectedGlyph(SmartyPantType type, string expected)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(
            new SmartyPant
            {
                Type = type
            });

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo(expected);
    }

    // The parser has already taken the characters out of the text by the time a kind reaches the
    // renderer, so one with no glyph is text gone from the document: <<x>> used to print as "x". A
    // kind Markdig adds later has to be given one here before it can ship.
    [Test]
    public async Task EveryKindHasAGlyph()
    {
        foreach (var type in Enum.GetValues<SmartyPantType>())
        {
            var renderer = RendererHarness.BuildRenderer();
            renderer.Render(
                new SmartyPant
                {
                    Type = type
                });

            await Assert.That(renderer.Top.CurrentRuns.Sum(_ => _.InnerText.Length)).IsGreaterThan(0);
        }
    }

    [Test]
    [Arguments("He said <<bonjour>> to all", "He said \u00ABbonjour\u00BB to all")]
    [Arguments("'single' and \"double\"", "\u2018single\u2019 and \u201Cdouble\u201D")]
    [Arguments("a -- b --- c", "a \u2013 b \u2014 c")]
    [Arguments("and so on...", "and so on\u2026")]
    // Not a quote on either side, so the parser leaves it as the text it is.
    [Arguments("x << 2 and y >> 3", "x << 2 and y >> 3")]
    public async Task ParsedTextRendersItsGlyphs(string markdown, string expected)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo(expected);
        await Assert.That(paragraph.Elements<Run>().Count()).IsEqualTo(1);
    }
}
