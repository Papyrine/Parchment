class EmphasisInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, EmphasisInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, EmphasisInline inline)
    {
        var before = renderer.Top.CurrentRuns.Count;
        renderer.WriteChildren(inline);

        // A {.StyleName} on emphasis ("**text**{.Lead}") binds to this inline. Nothing read it
        // before, so it was dropped without a word — the run kept the emphasis and lost the style.
        var styleId = MarkdownStyle.Resolve(inline);

        var top = renderer.Top;
        for (var i = before; i < top.CurrentRuns.Count; i++)
        {
            if (top.CurrentRuns[i] is Run run)
            {
                ApplyStyle(run, inline.DelimiterChar, inline.DelimiterCount);
                if (styleId != null)
                {
                    // Assigned through the typed property so rStyle lands first in the rPr sequence.
                    run.RunProperties!.RunStyle = new()
                    {
                        Val = styleId
                    };
                }
            }
        }
    }

    static void ApplyStyle(Run run, char delimiter, int count)
    {
        run.RunProperties ??= new();
        switch (delimiter)
        {
            case '*':
            case '_':
                if (count >= 2)
                {
                    run.RunProperties.Append(new Bold());
                }
                else
                {
                    run.RunProperties.Append(new Italic());
                }

                break;
            case '~':
                if (count >= 2)
                {
                    run.RunProperties.Append(new Strike());
                }
                else
                {
                    run.RunProperties.Append(
                        new VerticalTextAlignment
                        {
                            Val = VerticalPositionValues.Subscript
                        });
                }

                break;
            case '^':
                run.RunProperties.Append(
                    new VerticalTextAlignment
                    {
                        Val = VerticalPositionValues.Superscript
                    });
                break;
            case '+':
                run.RunProperties.Append(
                    new Underline
                    {
                        Val = UnderlineValues.Single
                    });
                break;
            case '=':
                run.RunProperties.Append(
                    new Highlight
                    {
                        Val = HighlightColorValues.Yellow
                    });
                break;
        }
    }
}
