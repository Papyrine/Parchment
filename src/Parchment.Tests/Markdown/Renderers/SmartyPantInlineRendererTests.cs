public class SmartyPantInlineRendererTests
{
    [Test]
    [Arguments(SmartyPantType.LeftQuote, "\u2018")]
    [Arguments(SmartyPantType.RightQuote, "\u2019")]
    [Arguments(SmartyPantType.LeftDoubleQuote, "\u201C")]
    [Arguments(SmartyPantType.RightDoubleQuote, "\u201D")]
    [Arguments(SmartyPantType.Dash2, "\u2013")]
    [Arguments(SmartyPantType.Dash3, "\u2014")]
    [Arguments(SmartyPantType.Ellipsis, "\u2026")]
    public async Task EmitsExpectedGlyph(SmartyPantType type, string expected)
    {
        var renderer = RendererHarness.BuildRenderer();
        renderer.Render(new SmartyPant { Type = type });

        var run = (Run)renderer.Top.CurrentRuns.Single();
        await Assert.That(run.GetFirstChild<Text>()!.Text).IsEqualTo(expected);
    }
}
