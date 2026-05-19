static class MarkdigPipeline
{
    public static MarkdownPipeline Pipeline { get; } = new MarkdownPipelineBuilder()
        .UseEmphasisExtras()
        .UseGridTables()
        .UsePipeTables(new() { InferColumnWidthsFromSeparator = true })
        .UseAutoLinks()
        .UseListExtras()
        .UseSmartyPants()
        .UseGenericAttributes()
        .Build();
}
