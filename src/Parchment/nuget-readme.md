# Parchment

Parchment is a Word (.docx) generation library that combines a .NET data model with either a docx template (token replacement) or a markdown template (full content rendering), driven by [liquid](https://shopify.github.io/liquid/) via [Fluid](https://github.com/sebastienros/fluid), [Markdig](https://github.com/xoofx/markdig), and [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml).

A template belongs to its model type and is rendered by it — one template per model, no names to repeat.

## Source generator (recommended)

Decorate the model class itself with `[ParchmentModel]` and Parchment's source generator finds the template by convention — the `AdditionalFiles` entry named after the type (`Invoice.docx` or `Report.md`) — validates its tokens against the model at compile time, and embeds it into the generated source. A module initializer registers it when the assembly loads, so rendering is the whole API:

```cs
[ParchmentModel]
public partial class Invoice
{
    public string Number { get; set; } = "";
    // ...
}

var store = new TemplateStore();

using var stream = new MemoryStream();
await store.Render(invoice, stream);
```

A markdown template's style source is found by convention too: `Report.dotx`, or the nearest `parchment.dotx` up the directory tree — embedded alongside the template.

## Registering by hand

For templates the generator cannot see — content produced at runtime, or a model that cannot be made `partial` — register against the model type directly.

A docx template:

```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<Invoice>("invoice-template.docx");

using var stream = new MemoryStream();
await store.Render(invoice, stream);
```

A markdown template, with an optional style source — a `Stream` over a Word template (`.dotx`) whose styles the output inherits:

```cs
store.RegisterMarkdownTemplate<Report>(markdownSource, styleSource);
```

`Render` writes to a stream the caller supplies. To write straight to disk, use `RenderToFile`:

```cs
await store.RenderToFile(invoice, "out.docx");
```

## Template content

The template may include:

- Substitution tokens: `{{ Customer.Name }}`
- Paragraph-scope loops: `{% for line in Lines %}` … `{% endfor %}`
- Table-row-scope loops: put `{% for line in Lines %}` on its own in one row and `{% endfor %}` on its own in another
- Conditionals: `{% if Customer.IsPreferred %}` … `{% endif %}`

Members are resolved against the model by name. There is no snake-case translation layer, so use the property names as declared.

See the [readme](https://github.com/Papyrine/Parchment#readme) for full documentation.
