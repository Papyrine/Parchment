class ListBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, ListBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, ListBlock listBlock)
    {
        var numId = listBlock.IsOrdered
            ? renderer.Numbering.CreateOrderedNumbering(MapOrderedFormat(listBlock))
            : renderer.Numbering.CreateBulletNumbering();

        foreach (var item in listBlock)
        {
            if (item is not ListItemBlock itemBlock)
            {
                continue;
            }

            foreach (var child in itemBlock)
            {
                if (child is not LeafBlock leaf)
                {
                    renderer.Render(child);
                    continue;
                }

                var properties = new ParagraphProperties
                {
                    ParagraphStyleId = new()
                    {
                        Val = "ListParagraph"
                    },
                    NumberingProperties = new(
                        new NumberingLevelReference
                        {
                            Val = 0
                        },
                        new NumberingId
                        {
                            Val = numId
                        }),
                    ContextualSpacing = new()
                };
                renderer.WriteLeafInline(leaf);
                renderer.FlushParagraph(properties);
            }
        }
    }

    static NumberFormatValues MapOrderedFormat(ListBlock listBlock) =>
        listBlock.BulletType switch
        {
            '1' => NumberFormatValues.Decimal,
            'a' => NumberFormatValues.LowerLetter,
            'A' => NumberFormatValues.UpperLetter,
            'i' => NumberFormatValues.LowerRoman,
            'I' => NumberFormatValues.UpperRoman,
            _ => NumberFormatValues.Decimal
        };
}
