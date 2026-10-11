class AutolinkInlineRenderer :
    MarkdownObjectRenderer<OpenXmlMarkdownRenderer, AutolinkInline>
{
    protected override void Write(OpenXmlMarkdownRenderer renderer, AutolinkInline inline)
    {
        var url = inline.IsEmail ? $"mailto:{inline.Url}" : inline.Url;
        var relId = RelationshipIds.Hyperlink(renderer.MainPart, url);
        var run = new Run(
            new RunProperties(
                new RunStyle
                {
                    Val = MarkdownStyle.Resolve(inline) ?? "Hyperlink"
                }),
            new Text(inline.Url)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        renderer.AddRun(
            new Hyperlink(run)
            {
                Id = relId
            });
    }
}
