# Parchment

Parchment is a Word (.docx) generation library that combines a .NET data model with either a docx template (token replacement) or a markdown template (full content rendering), driven by [liquid](https://shopify.github.io/liquid/) via [Fluid](https://github.com/sebastienros/fluid), [Markdig](https://github.com/xoofx/markdig), and [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml).

## Docx template flow

```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<Invoice>("invoice", "invoice-template.docx");

using var stream = new MemoryStream();
await store.Render("invoice", SampleData.Invoice(), stream);
```

`Render` writes to a stream the caller supplies. To write straight to disk, use `RenderToFile`:

```cs
await store.RenderToFile("invoice", SampleData.Invoice(), "out.docx");
```

The template may include:

- Substitution tokens: `{{ Customer.Name }}`
- Paragraph-scope loops: `{% for line in Lines %}` … `{% endfor %}`
- Table-row-scope loops: put `{% for line in Lines %}` on its own in one row and `{% endfor %}` on its own in another
- Conditionals: `{% if Customer.IsPreferred %}` … `{% endif %}`

Members are resolved against the model by name. There is no snake-case translation layer, so use the property names as declared.

## Markdown template flow

The style source is a `Stream` over a docx whose styles the output inherits. It is optional — omit it for a blank default.

```cs
var store = new TemplateStore();
store.RegisterMarkdownTemplate<Report>("report", markdownSource, styleSource);

using var stream = new MemoryStream();
await store.Render("report", reportModel, stream);
```

## Source generator

Decorate the model class itself with `[ParchmentModel]` and Parchment's source generator validates the template tokens against it at compile time. Both `.docx` and `.md` templates are supported.

```cs
[ParchmentModel("Templates/invoice.docx")]
public partial class Invoice
{
    public string Number { get; set; } = "";
    // ...
}

[ParchmentModel("Templates/report.md")]
public partial class Report
{
    public string Title { get; set; } = "";
    // ...
}
```

See the [readme](https://github.com/Papyrine/Parchment#readme) for full documentation.
