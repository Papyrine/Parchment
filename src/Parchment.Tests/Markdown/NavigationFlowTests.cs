// End-to-end cover for the navigation features: what the renderer builds has to survive template
// binding, style-source cloning and packaging to reach Word as a working field.
public class NavigationFlowTests
{
    public class ReportModel
    {
        public required IReadOnlyList<Section> Sections { get; init; }
    }

    public class Section
    {
        public required string Anchor { get; init; }
        public required string Title { get; init; }
        public required string Body { get; init; }
    }

    [Test]
    public async Task ContentsAndPageReferences()
    {
        var markdown =
            """
            Contents {.TOCHeading}

            [TOC]{levels=1}

            # Summary

            | Section | Page |
            | --- | --- |
            {% for section in Sections -%}
            | [{{ section.Title }}](#{{ section.Anchor }}) | [](#{{ section.Anchor }}) |
            {% endfor %}
            {% for section in Sections %}
            # {{ section.Title }} {#{{ section.Anchor }}}

            {{ section.Body }}
            {% endfor %}
            """;

        using var styleSource = DocxTemplateBuilder.Build();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportModel>(markdown, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new ReportModel
            {
                Sections =
                [
                    new()
                    {
                        Anchor = "delivery",
                        Title = "Delivery",
                        Body = "Where the work landed."
                    },
                    new()
                    {
                        Anchor = "risks",
                        Title = "Risks",
                        Body = "What could still go wrong."
                    }
                ]
            },
            stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
