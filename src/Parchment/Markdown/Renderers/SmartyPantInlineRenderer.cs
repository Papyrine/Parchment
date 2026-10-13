using Markdig.Extensions.SmartyPants;

class SmartyPantInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, SmartyPant>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, SmartyPant inline) =>
        renderer.AddRun(
            new Run(
                new Text(Glyph(inline.Type))
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));

    // Every kind, because the parser has already taken the characters out of the text: one with no
    // glyph here is text missing from the document, which is how <<x>> came to print as "x". The two
    // straight quotes are resolved into a left or a right one before parsing finishes, so neither
    // reaches a render today; they map to themselves rather than to nothing. A kind Markdig adds
    // later is refused rather than dropped, and SmartyPantInlineRendererTests fails first.
    internal static string Glyph(SmartyPantType type) =>
        type switch
        {
            SmartyPantType.Quote => "'",
            SmartyPantType.LeftQuote => "\u2018",
            SmartyPantType.RightQuote => "\u2019",
            SmartyPantType.DoubleQuote => "\"",
            SmartyPantType.LeftDoubleQuote => "\u201C",
            SmartyPantType.RightDoubleQuote => "\u201D",
            SmartyPantType.LeftAngleQuote => "\u00AB",
            SmartyPantType.RightAngleQuote => "\u00BB",
            SmartyPantType.Dash2 => "\u2013",
            SmartyPantType.Dash3 => "\u2014",
            SmartyPantType.Ellipsis => "\u2026",
            _ => throw new InvalidOperationException($"Unknown smarty pants kind {type}")
        };
}
