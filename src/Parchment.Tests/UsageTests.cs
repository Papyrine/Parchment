public class UsageTests
{
    [Test]
    public async Task Substitution()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: SubstitutionInput
            Invoice {{ Number }}

            Customer: {{ Customer.Name }}

            Total: {{ Total }} {{ Currency }}
            // end-snippet
            """);

        #region Substitution

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task SubstitutionFromNonMemoryStream()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Invoice {{ Number }}

            Customer: {{ Customer.Name }}

            Total: {{ Total }} {{ Currency }}
            """);

        // Wrap in BufferedStream so the registration hits the non-MemoryStream path
        using var wrapped = new BufferedStream(template);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(wrapped);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task Markdown()
    {
        var markdownSource =
            """
            <!-- begin-snippet: MarkdownTemplate(lang=handlebars) -->
            # {{ Report.Title }}

            *Prepared by **{{ Report.Author }}** on {{ Report.Date }}*

            ## Summary

            {{ Report.Summary }}

            ## Findings

            | Area | Status | Owner |
            | --- | --- | --- |
            {% for finding in Report.Findings -%}
            | {{ finding.Area }} | {{ finding.Status }} | {{ finding.Owner }} |
            {% endfor %}

            ## Action items

            {% for item in Report.Actions %}
            1. **{{ item.Title }}** — {{ item.Detail }}
            {% endfor %}

            {% if Report.HasRisks %}
            > ⚠ Outstanding risks remain. See appendix for mitigation plan.
            {% else %}
            > No outstanding risks.
            {% endif %}
            <!-- end-snippet -->
            """;

        #region Markdown

        using var brandDocx = DocxTemplateBuilder.Build();
        var reportModel = SampleData.Report();

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(
            markdownSource,
            styleSource: brandDocx);

        using var stream = new MemoryStream();
        await store.Render(reportModel, stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task MarkdownFromNonMemoryStream()
    {
        var markdownSource =
            """
            # {{ Report.Title }}

            {{ Report.Summary }}
            """;

        using var styleSource = DocxTemplateBuilder.Build();
        // Wrap in BufferedStream so the registration hits the non-MemoryStream path
        using var wrapped = new BufferedStream(styleSource);

        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<ReportContext>(
            markdownSource,
            styleSource: wrapped);

        using var stream = new MemoryStream();
        await store.Render(SampleData.Report(), stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }
}
