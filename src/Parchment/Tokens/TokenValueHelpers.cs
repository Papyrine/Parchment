namespace Parchment;

public static class TokenValueHelpers
{
    public static TokenValue BulletList(IEnumerable<string> items) =>
        new OpenXmlToken(context =>
        {
            var numId = context.CreateBulletNumbering();
            var pStyle = ResolveListPStyle(context);
            return items.Select(item => BuildListParagraph(item, numId, 0, pStyle));
        });

    public static TokenValue NumberedList(IEnumerable<string> items, NumberFormatValues? format = null) =>
        new OpenXmlToken(context =>
        {
            var numId = context.CreateOrderedNumbering(format ?? NumberFormatValues.Decimal);
            var pStyle = ResolveListPStyle(context);
            return items.Select(item => BuildListParagraph(item, numId, 0, pStyle));
        });

    // The bullet's indentation comes from the abstractNum level (Indentation Left+Hanging in
    // WordNumberingState.BuildBulletLevel), not from ListParagraph's own pPr — so adopting the
    // host's pStyle keeps the bullets visually correct while picking up the surrounding font.
    static string ResolveListPStyle(IOpenXmlContext context)
    {
        var hostStyle = context.HostParagraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrEmpty(hostStyle))
        {
            return "ListParagraph";
        }

        return hostStyle;
    }

    static Paragraph BuildListParagraph(string text, int numId, int ilvl, string pStyle)
    {
        var paragraph = new Paragraph();
        var props = new ParagraphProperties
        {
            ParagraphStyleId = new()
            {
                Val = pStyle
            },
            NumberingProperties = new(
                new NumberingLevelReference
                {
                    Val = ilvl
                },
                new NumberingId
                {
                    Val = numId
                }),
            ContextualSpacing = new()
        };
        paragraph.Append(props);
        paragraph.Append(new Run(
            new Text(XmlCharSanitizer.Strip(text))
            {
                Space = SpaceProcessingModeValues.Preserve
            }));
        return paragraph;
    }
}
