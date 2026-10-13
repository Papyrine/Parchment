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
                Apply(run, inline, styleId);
                continue;
            }

            // A link arrives as one Hyperlink holding the runs of its text, so "**[text](url)**"
            // puts no Run of its own between the delimiters and the emphasis has to be carried
            // inside. The {.StyleName} goes with it: a run takes a single rStyle, and the one on
            // emphasis wins over a link's whichever of the two encloses the other.
            foreach (var inner in element.Descendants<Run>())
            {
                Apply(inner, inline, styleId);
            }
        }
    }

    // One run carries one character style, and the inlines around it can ask for several. Which
    // wins goes by what asked, not by nesting order: emphasis, then a classed span, then inline
    // code, then the link ("Hyperlink" unless it names its own). Emphasis is the last to reach a
    // run - it is applied once its children are written - so it replaces whatever is there;
    // HtmlInlineRenderer and LinkInlineRenderer hold up the other end.
    static void Apply(Run run, EmphasisInline inline, string? styleId)
    {
        ApplyStyle(run, inline.DelimiterChar, inline.DelimiterCount);
        if (styleId == null)
        {
            return;
        }

        // Assigned through the typed property so rStyle lands first in the rPr sequence.
        run.RunProperties!.RunStyle = new()
        {
            Val = styleId
        };
    }

    // Emphasis nests, and the outer one applies to runs the inner already reached: "*_text_*" walks
    // the same run twice. Every property here is one the schema lets a run carry once, so the second
    // pass leaves what the first put there rather than adding a duplicate beside it. Each goes in
    // through its typed property, which places it by the rPr sequence rather than last: a run inside
    // a link already holds an rStyle, and the inner emphasis is not always the earlier element -
    // "**_text_**" reaches the run italic first and bold second.
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
