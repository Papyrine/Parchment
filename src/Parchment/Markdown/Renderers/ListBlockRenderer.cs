class ListBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, ListBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, ListBlock listBlock)
    {
        var numId = listBlock.IsOrdered
            ? renderer.Numbering.CreateOrderedNumbering(MapOrderedFormat(listBlock))
            : renderer.Numbering.CreateBulletNumbering();
        var ilvl = ResolveIlvl(listBlock);

        foreach (var item in listBlock)
        {
            if (item is not ListItemBlock itemBlock)
            {
                continue;
            }

            foreach (var child in itemBlock)
            {
                if (child is HtmlBlock or CodeBlock or ThematicBreakBlock || child is not LeafBlock leaf)
                {
                    renderer.PushIndent(480);
                    try
                    {
                        renderer.Render(child);
                    }
                    finally
                    {
                        renderer.PopIndent();
                    }
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
                            Val = ilvl
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

    /// <summary>
    /// Nesting depth of a list, which is the <c>ilvl</c> its items belong at.
    /// </summary>
    /// <remarks>
    /// A nested list arrives as a fresh Write with its own numbering, so the depth has to come from
    /// the markdown tree rather than from renderer state. Every level used to be emitted at
    /// <c>ilvl=0</c>, which flattened nesting: the marker never changed and the indentation, which
    /// the abstractNum's level supplies, never applied.
    /// </remarks>
    static int ResolveIlvl(ListBlock listBlock)
    {
        var depth = 0;
        for (var parent = listBlock.Parent; parent != null; parent = parent.Parent)
        {
            if (parent is ListBlock)
            {
                depth++;
            }
        }

        return Math.Min(depth, WordNumberingState.MaxIlvl);
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
