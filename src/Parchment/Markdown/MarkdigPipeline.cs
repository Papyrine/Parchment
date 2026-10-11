using Markdig.Parsers.Inlines;

static class MarkdigPipeline
{
    public static MarkdownPipeline Pipeline { get; } = Build();

    static MarkdownPipeline Build()
    {
        var builder = new MarkdownPipelineBuilder()
            .UseEmphasisExtras()
            .UseGridTables()
            .UsePipeTables(new() { InferColumnWidthsFromSeparator = true })
            .UseListExtras()
            .UseSmartyPants()
            .UseGenericAttributes();

        // What UseAutoLinks registers, with the parser swapped for the one that understands the
        // escapes in a bound value.
        builder.InlineParsers.InsertBefore<EmphasisInlineParser>(new BoundAutoLinkParser(new()));
        return builder.Build();
    }
}
