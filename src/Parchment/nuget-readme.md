# Parchment

Parchment is a Word (.docx) generation library that combines a .NET data model with either a docx template (token replacement) or a markdown template (full content rendering), driven by [liquid](https://shopify.github.io/liquid/) via [Fluid](https://github.com/sebastienros/fluid), [Markdig](https://github.com/xoofx/markdig), and [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml).

A template is registered against its model type and rendered by it — one template per model, no names to repeat.

## Docx template flow

```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<Invoice>("invoice-template.docx");

using var stream = new MemoryStream();
await store.Render(SampleData.Invoice(), stream);
```

`Render` writes to a stream the caller supplies. To write straight to disk, use `RenderToFile`:

```cs
await store.RenderToFile(SampleData.Invoice(), "out.docx");
```

The template may include:

- Substitution tokens: `{{ Customer.Name }}`
- Paragraph-scope loops: `{% for line in Lines %}` … `{% endfor %}`
- Table-row-scope loops: put `{% for line in Lines %}` on its own in one row and `{% endfor %}` on its own in another
- Conditionals: `{% if Customer.IsPreferred %}` … `{% endif %}`

Members are resolved against the model by name. There is no snake-case translation layer, so use the property names as declared.

## Markdown template flow

The style source is a `Stream` over a Word template (`.dotx`) whose styles the output inherits. It is optional — omit it for a blank default.

```cs
var store = new TemplateStore();
store.RegisterMarkdownTemplate<Report>(markdownSource, styleSource);

using var stream = new MemoryStream();
await store.Render(reportModel, stream);
```

## Source generator

Decorate the model class itself with `[ParchmentModel]` and Parchment's source generator finds the template by convention — the `AdditionalFiles` entry named after the type (`Invoice.docx` or `Report.md`) — validates its tokens against the model at compile time, and embeds it into the generated source. A module initializer registers it when the assembly loads, so rendering needs no setup:

```cs
[ParchmentModel]
public partial class Invoice
{
    public string Number { get; set; } = "";
    // ...
}

var store = new TemplateStore();
await store.Render(invoice, stream);
```

A markdown template's style source is found by convention too: `Report.dotx`, or the nearest `parchment.dotx` up the directory tree.

See the [readme](https://github.com/Papyrine/Parchment#readme) for full documentation.
