class ListBlockRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, ListBlock>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, ListBlock listBlock)
    {
        var ilvl = ResolveIlvl(listBlock);
        var numId = listBlock.IsOrdered
            ? renderer.Numbering.CreateOrderedNumbering(MapOrderedFormat(listBlock), ilvl, ResolveStart(listBlock))
            : renderer.Numbering.CreateBulletNumbering();
        var listStyleId = MarkdownStyle.Resolve(listBlock);

        foreach (var item in listBlock)
        {
            if (item is not ListItemBlock itemBlock)
            {
                continue;
            }

            // Only the item's first paragraph takes the marker. Numbering every paragraph gave a
            // continuation paragraph a bullet of its own and made an ordered list count it as the
            // next item.
            var numbered = false;
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

                // {.StyleName} on an item ("- text{.Caption}") attaches to the item's own leaf;
                // one before the list attaches to the ListBlock and covers every item. The item
                // wins where both are present.
                var heading = leaf as HeadingBlock;
                var properties = new ParagraphProperties();
                if (heading == null)
                {
                    properties.ParagraphStyleId = new()
                    {
                        Val = MarkdownStyle.Resolve(leaf) ?? listStyleId ?? "ListParagraph"
                    };
                    properties.ContextualSpacing = new();
                }

                if (numbered)
                {
                    // Without a numPr none of the level's indentation applies, so the paragraph
                    // states the level's text indent itself to sit under the item's text.
                    properties.Indentation = new()
                    {
                        Left = WordNumberingState.TextIndent(ilvl).ToString(CultureInfo.InvariantCulture)
                    };
                }
                else
                {
                    properties.NumberingProperties = new(
                        new NumberingLevelReference
                        {
                            Val = ilvl
                        },
                        new NumberingId
                        {
                            Val = numId
                        });
                    numbered = true;
                }

                if (heading != null)
                {
                    // A heading takes the marker and the indent like any other item text, but
                    // under its own style rather than the list's, so it is still a heading.
                    HeadingBlockRenderer.Write(renderer, heading, properties);
                    continue;
                }

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

    /// <summary>
    /// The number an ordered list starts at, so "5. five" is numbered 5 rather than 1.
    /// </summary>
    /// <remarks>
    /// <see cref="ListBlock.OrderedStart"/> is a string: digits for a decimal list, and for the
    /// alpha and roman list-extras either the authored marker ("c", "iv") or its value. Both
    /// shapes are read. Anything unreadable falls back to 1, which is what every list used to get.
    /// </remarks>
    static int ResolveStart(ListBlock listBlock)
    {
        var start = listBlock.OrderedStart;
        if (string.IsNullOrEmpty(start))
        {
            return 1;
        }

        if (int.TryParse(start, NumberStyles.None, CultureInfo.InvariantCulture, out var number))
        {
            return number;
        }

        if (listBlock.BulletType is 'a' or 'A')
        {
            return ParseAlpha(start);
        }

        if (listBlock.BulletType is 'i' or 'I')
        {
            return ParseRoman(start);
        }

        return 1;
    }

    static int ParseAlpha(string value)
    {
        if (value.Length == 1 &&
            char.IsAsciiLetter(value[0]))
        {
            return char.ToLowerInvariant(value[0]) - 'a' + 1;
        }

        return 1;
    }

    static int ParseRoman(string value)
    {
        var total = 0;
        var previous = 0;
        for (var index = value.Length - 1; index >= 0; index--)
        {
            var digit = RomanDigit(value[index]);
            if (digit == 0)
            {
                return 1;
            }

            if (digit < previous)
            {
                total -= digit;
                continue;
            }

            total += digit;
            previous = digit;
        }

        return total;
    }

    static int RomanDigit(char value) =>
        char.ToLowerInvariant(value) switch
        {
            'i' => 1,
            'v' => 5,
            'x' => 10,
            'l' => 50,
            'c' => 100,
            'd' => 500,
            'm' => 1000,
            _ => 0
        };

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
