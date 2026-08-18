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
        var properties = run.RunProperties ??= new();
        switch (delimiter)
        {
            case '*':
            case '_':
                if (count >= 2)
                {
                    Set(properties, new Bold());
                }
                else
                {
                    Set(properties, new Italic());
                }

                break;
            case '~':
                if (count >= 2)
                {
                    Set(properties, new Strike());
                }
                else
                {
                    Set(
                        properties,
                        new VerticalTextAlignment
                        {
                            Val = VerticalPositionValues.Subscript
                        });
                }

                break;
            case '^':
                Set(
                    properties,
                    new VerticalTextAlignment
                    {
                        Val = VerticalPositionValues.Superscript
                    });
                break;
            case '+':
                Set(
                    properties,
                    new Underline
                    {
                        Val = UnderlineValues.Single
                    });
                break;
            case '=':
                Set(
                    properties,
                    new Highlight
                    {
                        Val = HighlightColorValues.Yellow
                    });
                break;
        }
    }

    // Emphasis nests, and the outer one applies to runs the inner already reached: "*_text_*" walks
    // the same run twice. Every property here is one the schema lets a run carry once, so the second
    // pass leaves what the first put there rather than adding a duplicate beside it.
    static void Set<T>(RunProperties properties, T property)
        where T : OpenXmlElement
    {
        if (properties.GetFirstChild<T>() == null)
        {
            properties.Append(property);
        }
    }
}
