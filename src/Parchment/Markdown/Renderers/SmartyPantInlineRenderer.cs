using Markdig.Extensions.SmartyPants;

class SmartyPantInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, SmartyPant>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, SmartyPant inline)
    {
        var glyph = inline.Type switch
        {
            SmartyPantType.LeftQuote => "\u2018",
            SmartyPantType.RightQuote => "\u2019",
            SmartyPantType.LeftDoubleQuote => "\u201C",
            SmartyPantType.RightDoubleQuote => "\u201D",
            SmartyPantType.Dash2 => "\u2013",
            SmartyPantType.Dash3 => "\u2014",
            SmartyPantType.Ellipsis => "\u2026",
            _ => string.Empty
        };

        if (glyph.Length == 0)
        {
            return;
        }

        renderer.AddRun(
            new Run(
                new Text(glyph)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
    }
}
