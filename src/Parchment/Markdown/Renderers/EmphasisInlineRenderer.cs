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
            var element = top.CurrentRuns[i];
            if (element is Run run)
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

                continue;
            }

            // A link arrives as one Hyperlink holding the runs of its text, so "**[text](url)**"
            // puts no Run of its own between the delimiters and the emphasis has to be carried
            // inside. The {.StyleName} is not: a run takes a single rStyle, and replacing the one
            // the link put there would stop the text reading as a link.
            foreach (var inner in element.Descendants<Run>())
            {
                ApplyStyle(inner, inline.DelimiterChar, inline.DelimiterCount);
            }
        }
    }

    // Emphasis nests, and the outer one applies to runs the inner already reached: "*_text_*" walks
    // the same run twice. Every property here is one the schema lets a run carry once, so the second
    // pass leaves what the first put there rather than adding a duplicate beside it. Each goes in
    // through its typed property, which places it by the rPr sequence rather than last: a run inside
    // a link already holds an rStyle, and the inner emphasis is not always the earlier element.
    static void ApplyStyle(Run run, char delimiter, int count)
    {
        var properties = run.RunProperties ??= new();
        switch (delimiter)
        {
            case '*':
            case '_':
                if (count >= 2)
                {
                    properties.Bold ??= new();
                }
                else
                {
                    properties.Italic ??= new();
                }

                break;
            case '~':
                if (count >= 2)
                {
                    properties.Strike ??= new();
                }
                else
                {
                    properties.VerticalTextAlignment ??= new()
                    {
                        Val = VerticalPositionValues.Subscript
                    };
                }

                break;
            case '^':
                properties.VerticalTextAlignment ??= new()
                {
                    Val = VerticalPositionValues.Superscript
                };
                break;
            case '+':
                properties.Underline ??= new()
                {
                    Val = UnderlineValues.Single
                };
                break;
            case '=':
                properties.Highlight ??= new()
                {
                    Val = HighlightColorValues.Yellow
                };
                break;
        }
    }
}
