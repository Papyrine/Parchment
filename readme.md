# <img src="/src/icon.png" height="30px"> Parchment

[![Build status](https://img.shields.io/appveyor/build/SimonCropp/Parchment)](https://ci.appveyor.com/project/SimonCropp/Parchment)
[![NuGet Status](https://img.shields.io/nuget/v/Parchment.svg?label=Parchment)](https://www.nuget.org/packages/Parchment/)

Parchment is a Word (.docx) document generation library with two complementary rendering modes. It combines a .NET data model with either a docx template (token replacement, loops, conditionals) or a markdown template (full content rendering), both driven by [liquid](https://shopify.github.io/liquid/) via [Fluid](https://github.com/sebastienros/fluid). Markdown is parsed with [Markdig](https://github.com/xoofx/markdig); HTML chunks are converted via [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml).

**See [Milestones](../../milestones?state=closed) for release notes.**


## Open Source Maintenance Fee

This project participates in the [Open Source Maintenance Fee](https://opensourcemaintenancefee.org). The source code is freely available under the terms of the [license](license.txt). To support sustainable maintenance, use of the project's official binary releases in revenue-generating activities and all government agencies requires adherence to the [Open Source Maintenance Fee EULA](OsmfEula.txt). The fee is paid by [sponsoring Papyrine](https://github.com/sponsors/Papyrine).

This project uses [SponsorCheck](https://github.com/SimonCropp/SponsorCheck) to surface a build-time reminder in consuming projects that are not yet sponsoring.


## NuGet package

[Parchment](https://www.nuget.org/packages/Parchment/)


## Two modes

Parchment supports two complementary template formats:

1. **Docx template** — start from a hand-crafted Word document with `{{ field }}` substitution tokens and `{% for %}` / `{% if %}` block tags. Output preserves every detail of the source document.
2. **Markdown template** — start from a `.md` file with full liquid support. Markdown is parsed by Markdig and rendered into a target docx. Optionally provide a style-source `.docx` whose styles, headers, footers, and section properties are inherited.


## Docx template


### Input docx content

<!-- snippet: SubstitutionInput -->
<a id='snippet-SubstitutionInput'></a>
```cs
Invoice {{ Number }}

Customer: {{ Customer.Name }}

Total: {{ Total }} {{ Currency }}
```
<sup><a href='/src/Parchment.Tests/UsageTests.cs#L8-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-SubstitutionInput' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Render

<!-- snippet: Substitution -->
<a id='snippet-Substitution'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<Invoice>("substitution", template);

using var stream = new MemoryStream();
await store.Render("substitution", SampleData.Invoice(), stream);
```
<sup><a href='/src/Parchment.Tests/UsageTests.cs#L17-L25' title='Snippet source file'>snippet source</a> | <a href='#snippet-Substitution' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Token naming

Tokens match model members case-insensitively via Fluid's default member access strategy, but Parchment uses PascalCase by convention (`{{ Customer.Name }}`), not snake_case. There is no snake-case translation layer — stick with the property name as declared on the model.

Tokens in `MainDocumentPart`, every `HeaderPart` / `FooterPart`, `FootnotesPart`, and `EndnotesPart` are all scanned and substituted, so page headers, footers, and footnotes can use `{{ ... }}` and `{% ... %}` exactly like body paragraphs.


### Loops

A `{% for %}` paragraph and its matching `{% endfor %}` repeat the intervening content once per item.

```
{% for line in Lines %}
- {{ line.Description }}: {{ line.Quantity }} x {{ line.UnitPrice }}
{% endfor %}
```

The body isn't restricted to paragraphs. Anything between the open and close paragraphs at the same nesting level — additional paragraphs, `<w:tbl>` tables, images, section breaks — is captured and cloned per iteration. This makes per-group "heading + table" layouts straightforward: place a styled heading paragraph and a hand-crafted Word table between `{% for %}` and `{% endfor %}`, and reference the loop variable inside the table cells.

```
{% for dept in Departments %}
[heading paragraph: {{ dept.Name }}]
[Word table with cells {{ dept.X }}, {{ dept.Y }}, ...]
{% endfor %}
```

The open and close tags must be siblings (same parent element). A `{% for %}` paragraph in the body and an `{% endfor %}` paragraph inside a table cell will not pair up.

### Conditionals

```
{% if Customer.IsPreferred %}
Preferred customer: {{ Customer.Name }}
{% else %}
Regular customer: {{ Customer.Name }}
{% endif %}
```


### Nested loops and conditionals

Loops and conditionals can be nested to arbitrary depth. The outer loop variable stays in scope for inner constructs.

```
{% for group in Groups %}
{{ group.Name }}
{% for item in group.Items %}
- {{ item }}
{% endfor %}
{% endfor %}
```


### Token override hatches

Declare a model property as `TokenValue` and return one of:

- A plain `string` (or any value Fluid stringifies) — plain text substitution. This is the default when the property is typed as `string` directly; assigning a string to a `TokenValue` property goes through the same path via the implicit conversion.
- `new MarkdownToken(string)` — the value is rendered as markdown via Markdig and spliced into the host paragraph.
- `new HtmlToken(string)` — the value is converted from HTML (via [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml)) and spliced into the host paragraph.
- `new OpenXmlToken(Func<IOpenXmlContext, IEnumerable<OpenXmlElement>>)` — the callback emits raw OpenXML elements. Useful for rich tables, generated charts, custom-styled lists.
- `new MutateToken(Action<Paragraph, IOpenXmlContext>)` — the callback receives the host paragraph and mutates it in place. The token text is cleared before the callback runs. Useful for adding runs with custom formatting, injecting bookmarks, or tweaking paragraph properties while preserving the original paragraph.

#### List filters

`bullet_list` and `numbered_list` render an `IEnumerable<string>` property as a real Word list (`<w:numPr>` with a proper numbering definition), not as literal text. The token must sit alone in its paragraph — the host paragraph is replaced with one `<w:p>` per item.

Content:

<!-- snippet: BulletListFilterContent -->
<a id='snippet-BulletListFilterContent'></a>
```cs
using var template = DocxTemplateBuilder.Build(
    """
    Tags:

    {{ Tags | bullet_list }}
    """);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L267-L276' title='Snippet source file'>snippet source</a> | <a href='#snippet-BulletListFilterContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: BulletListFilterRender -->
<a id='snippet-BulletListFilterRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<Invoice>("bullet-filter", template);
using var stream = new MemoryStream();
await store.Render("bullet-filter", SampleData.Invoice(), stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L278-L285' title='Snippet source file'>snippet source</a> | <a href='#snippet-BulletListFilterRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`numbered_list` is identical in shape — swap the filter name to produce a decimal-numbered list instead of bullets.


#### `escape_xml` filter

Escapes `<`, `>`, `&`, `"`, and `'` in a string value. Useful when a token's value is spliced into a context where raw markup would be interpreted — e.g. feeding a user-supplied string into a `MarkdownToken` source that contains HTML blocks:

```
{{ UserSuppliedComment | escape_xml }}
```


#### Markdown property

Declare a model property as `TokenValue` and return `new MarkdownToken(...)` to inject rendered markdown at the substitution site:

Model:

<!-- snippet: MarkdownPropertyModel -->
<a id='snippet-MarkdownPropertyModel'></a>
```cs
public class NoteModel
{
    public required string Title;
    public required TokenValue Body;
}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L9-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownPropertyModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Content:

<!-- snippet: MarkdownPropertyContent -->
<a id='snippet-MarkdownPropertyContent'></a>
```cs
# {{ Title }}

{{ Body }}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L25-L29' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownPropertyContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: MarkdownPropertyRender -->
<a id='snippet-MarkdownPropertyRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<NoteModel>("markdown-hatch", template);
await store.Render(
    "markdown-hatch",
    new NoteModel
    {
        Title = "Weekly summary",
        Body = new MarkdownToken(
            """
            ## Highlights

            - Shipped the **new feature**
            - Closed _several_ bugs
            - Ran a code review

            > Stay the course
            """)
    },
    stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L32-L54' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownPropertyRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Markdown filter

Alternatively, use the `| markdown` filter on a plain `string` property:

Model:

<!-- snippet: MarkdownFilterModel -->
<a id='snippet-MarkdownFilterModel'></a>
```cs
public class ArticleModel
{
    public required string Heading;
    public required string Content;
}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L60-L68' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownFilterModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Content:

<!-- snippet: MarkdownFilterContent -->
<a id='snippet-MarkdownFilterContent'></a>
```cs
# {{ Heading }}

{{ Content | markdown }}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L76-L80' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownFilterContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: MarkdownFilterRender -->
<a id='snippet-MarkdownFilterRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<ArticleModel>("markdown-filter", template);
await store.Render(
    "markdown-filter",
    new ArticleModel
    {
        Heading = "Release notes",
        Content =
            """
            ### Bug fixes

            - Fixed crash on **empty input**
            - Resolved _timeout_ in batch mode
            """
    },
    stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L83-L102' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownFilterRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Both approaches produce the same structural replacement — the host paragraph is swapped with the rendered markdown elements when the token is the entire paragraph. If the token shares its paragraph with other text or with sibling tokens, the runtime falls back to inline splicing (single produced paragraph → its runs are extracted and merged into the host) or paragraph splitting (multiple produced blocks → host is split at the token offset and the produced blocks slot between the two halves). See [Inline-aware structural replacement](#inline-aware-structural-replacement) for the full rules.

**Markdown templates**: Neither `MarkdownToken` nor `| markdown` is needed when using `RegisterMarkdownTemplate`. The entire template is already markdown — a plain `string` property containing markdown syntax is interpolated into the source before Markdig parses it, so formatting works automatically:

Model:

<!-- snippet: MarkdownTemplatePropertyModel -->
<a id='snippet-MarkdownTemplatePropertyModel'></a>
```cs
public class BriefModel
{
    public required string Title;
    public required string Details;
}
```
<sup><a href='/src/Parchment.Tests/Markdown/MarkdownFlowTests.cs#L57-L65' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownTemplatePropertyModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Content:

<!-- snippet: MarkdownTemplatePropertyContent -->
<a id='snippet-MarkdownTemplatePropertyContent'></a>
```handlebars
# {{ Title }}

{{ Details }}
```
<sup><a href='/src/Parchment.Tests/Markdown/MarkdownFlowTests.cs#L73-L77' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownTemplatePropertyContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: MarkdownTemplatePropertyUsage -->
<a id='snippet-MarkdownTemplatePropertyUsage'></a>
```cs
var store = new TemplateStore();
store.RegisterMarkdownTemplate<BriefModel>(
    "brief",
    markdown,
    styleSource);

await store.Render(
    "brief",
    new BriefModel
    {
        Title = "Sprint recap",
        Details =
            """
            ## Done

            - Landed the **search** feature
            - Fixed _three_ regressions

            > Ship it.
            """
    },
    targetStream);
```
<sup><a href='/src/Parchment.Tests/Markdown/MarkdownFlowTests.cs#L82-L107' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownTemplatePropertyUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Html property

Declare a model property as `TokenValue` and return `new HtmlToken(...)` to convert HTML source to Word elements at the substitution site. Use this when the value arrives as HTML — typically from a CMS, a WYSIWYG editor, or any system that already speaks HTML:

Model:

<!-- snippet: HtmlPropertyModel -->
<a id='snippet-HtmlPropertyModel'></a>
```cs
public class PostModel
{
    public required string Title;
    public required TokenValue Body;
}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L159-L167' title='Snippet source file'>snippet source</a> | <a href='#snippet-HtmlPropertyModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Content:

<!-- snippet: HtmlPropertyContent -->
<a id='snippet-HtmlPropertyContent'></a>
```cs
# {{ Title }}

{{ Body }}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L175-L179' title='Snippet source file'>snippet source</a> | <a href='#snippet-HtmlPropertyContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: HtmlPropertyRender -->
<a id='snippet-HtmlPropertyRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<PostModel>("html-hatch", template);
await store.Render(
    "html-hatch",
    new PostModel
    {
        Title = "Welcome",
        Body = new HtmlToken(
            """
            <p>Welcome to the <b>weekly digest</b>.</p>
            <ul>
              <li>Search performance is up</li>
              <li>Two regressions closed</li>
            </ul>
            """)
    },
    stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L182-L202' title='Snippet source file'>snippet source</a> | <a href='#snippet-HtmlPropertyRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### OpenXml property

When the output can't be expressed as markdown or HTML — generated tables, charts, custom-styled callouts — use `new OpenXmlToken(...)` to emit raw OpenXML directly. The callback receives an `IOpenXmlContext` exposing numbering, image parts, and style lookups:

Model:

<!-- snippet: OpenXmlPropertyModel -->
<a id='snippet-OpenXmlPropertyModel'></a>
```cs
public class ReportModel
{
    public required string Title;
    public required TokenValue Callout;
}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L208-L216' title='Snippet source file'>snippet source</a> | <a href='#snippet-OpenXmlPropertyModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Content:

<!-- snippet: OpenXmlPropertyContent -->
<a id='snippet-OpenXmlPropertyContent'></a>
```cs
# {{ Title }}

{{ Callout }}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L224-L228' title='Snippet source file'>snippet source</a> | <a href='#snippet-OpenXmlPropertyContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: OpenXmlPropertyRender -->
<a id='snippet-OpenXmlPropertyRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<ReportModel>("openxml-hatch", template);
await store.Render(
    "openxml-hatch",
    new ReportModel
    {
        Title = "Status",
        Callout = new OpenXmlToken(_ =>
        [
            new Paragraph(
                new Run(
                    new RunProperties(
                        new Color
                        {
                            Val = "C00000"
                        },
                        new Bold()),
                    new Text("Critical: review required")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }))
        ])
    },
    stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L231-L258' title='Snippet source file'>snippet source</a> | <a href='#snippet-OpenXmlPropertyRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Mutate paragraph

Use `new MutateToken(...)` to receive the host paragraph and modify it in place. The token text is cleared before the callback runs, so the original paragraph structure (properties, styles) is preserved:

Model:

<!-- snippet: MutateModel -->
<a id='snippet-MutateModel'></a>
```cs
public class StyledModel
{
    public required string Label;
    public required TokenValue Highlight;
}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L108-L116' title='Snippet source file'>snippet source</a> | <a href='#snippet-MutateModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Content:

<!-- snippet: MutateContent -->
<a id='snippet-MutateContent'></a>
```cs
{{ Label }}

{{ Highlight }}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L124-L128' title='Snippet source file'>snippet source</a> | <a href='#snippet-MutateContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render:

<!-- snippet: MutateRender -->
<a id='snippet-MutateRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<StyledModel>("mutate", template);
await store.Render(
    "mutate",
    new StyledModel
    {
        Label = "Before",
        Highlight = new MutateToken((paragraph, _) =>
        {
            paragraph.Append(
                new Run(
                    new RunProperties(
                        new Bold()),
                    new Text("Custom content")
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }));
        })
    }, stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L131-L153' title='Snippet source file'>snippet source</a> | <a href='#snippet-MutateRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### Excelsior tables

Mark any collection property on the model with `[ExcelsiorTable]` and the matching `{{ ... }}` substitution is rendered as a fully-formatted Word table by [Excelsior](https://github.com/Papyrine/Excelsior) at render time. Headings, column ordering, formatting, null display, and custom render callbacks all come from Excelsior's `[Column]` attribute on the element type — the same configuration surface used for spreadsheets.

Mark the collection on the model:

<!-- snippet: ExcelsiorTableModel -->
<a id='snippet-ExcelsiorTableModel'></a>
```cs
public class Quote
{
    public required string Reference;

    [ExcelsiorTable]
    public required IReadOnlyList<QuoteLine> Lines;
}

public class QuoteLine
{
    [Column(Heading = "Item", Order = 1)]
    public required string Description;

    [Column(Heading = "Qty", Order = 2)]
    public required int Quantity;

    [Column(Order = 3, Format = "C0")]
    public required decimal UnitPrice;
}
```
<sup><a href='/src/Parchment.Tests/Docx/ExcelsiorTableTests.cs#L12-L34' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelsiorTableModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Drop a `{{ Lines }}` substitution into the template on its own line. The template:

![Template before render](/src/Parchment.Tests/Scenarios/excelsior-table/input.png)

Register and render normally:

<!-- snippet: ExcelsiorTableUsage -->
<a id='snippet-ExcelsiorTableUsage'></a>
```cs
var templatePath = Path.Combine(ScenarioPath("excelsior-table"), "input.docx");

var store = new TemplateStore();
store.RegisterDocxTemplate<Quote>("excelsior-quote", templatePath);

var model = new Quote
{
    Reference = "Q-2026-0042",
    Lines =
    [
        new()
        {
            Description = "Strategy workshop",
            Quantity = 2,
            UnitPrice = 4500m
        },
        new()
        {
            Description = "Implementation support",
            Quantity = 8,
            UnitPrice = 1750m
        },
        new()
        {
            Description = "Documentation review",
            Quantity = 1,
            UnitPrice = 950m
        }
    ]
};

using var stream = new MemoryStream();
await store.Render("excelsior-quote", model, stream);
```
<sup><a href='/src/Parchment.Tests/Docx/ExcelsiorTableTests.cs#L300-L336' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelsiorTableUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The rendered output:

![Rendered output](/src/Parchment.Tests/Scenarios/excelsior-table/output%23page01.verified.png)

Rules:

- The substitution must sit alone in its own paragraph — structural table replacement swaps the entire host paragraph, so surrounding text would be discarded. The runtime throws at registration (`ParchmentRegistrationException`) and the source generator emits `PARCH007` if this is violated.
- The substitution must be a plain member-access expression — filters (`{{ Lines | reverse }}`) and arithmetic are rejected because the Excelsior path walks the model object directly and bypasses Fluid evaluation. Diagnostic `PARCH008` covers this at compile time.
- Nested paths like `{{ Customer.Lines }}` work — the registration walks the model type recursively at build time, so `[ExcelsiorTable]` can sit on any reachable collection property.
- Currency and date formatting in the rendered table honor `Excelsior.ValueRenderer.Culture` (defaults to `CultureInfo.CurrentCulture`). Set it once in a module initializer to override the default locale.
- Loop-scoped tokens are **not** rendered as tables — the map is keyed on dotted paths from the root model, so `{{ dept.Lines }}` inside `{% for dept in Departments %}` falls through to plain Fluid output. For a per-group table inside a loop (or any table needing styling beyond the options below), build it directly with `OpenXmlToken` + Excelsior's `WordTableBuilder` — see [Loop-scoped or fully-custom tables](#loop-scoped-or-fully-custom-tables) below.

#### Branding the table from template styles

By default the rendered table uses Excelsior's default cell formatting, which falls back to the host document's `Normal` paragraph style. To drive the cell font, size, and spacing from named paragraph styles defined in the template, set `HeadingParagraphStyle` / `BodyParagraphStyle` on the attribute:

<!-- snippet: ExcelsiorTableParagraphStyles -->
<a id='snippet-ExcelsiorTableParagraphStyles'></a>
```cs
public class StyledQuote
{
    [ExcelsiorTable(HeadingParagraphStyle = "TBLHeading", BodyParagraphStyle = "TBLText")]
    public required IReadOnlyList<QuoteLine> Lines;
}
```
<sup><a href='/src/Parchment.Tests/Docx/ExcelsiorTableTests.cs#L108-L116' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelsiorTableParagraphStyles' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The style ids must exist in the template's styles part. Unlike inline run formatting, a paragraph style reaches **every** cell paragraph — including `IsHtml` and link cells — so the look is consistent across all content. (These map straight onto Excelsior's `WordTableBuilder.HeadingParagraphStyle`/`BodyParagraphStyle`.)

#### Loop-scoped or fully-custom tables

`[ExcelsiorTable]` only resolves for collections reachable from the root model — a token like `{{ dept.Lines }}` inside a `{% for dept in Departments %}` loop falls through to plain Fluid output. The same applies when a table needs styling or grouping beyond what the attribute exposes. In those cases, bypass the attribute and build the table directly with Excelsior's [`WordTableBuilder`](https://github.com/Papyrine/Excelsior#word-tables), returning it from an [`OpenXmlToken`](#custom-openxmltoken-programmatic-embedding) on a `TokenValue`-typed property:

<!-- snippet: ExcelsiorTableViaOpenXmlToken -->
<a id='snippet-ExcelsiorTableViaOpenXmlToken'></a>
```cs
public class GroupedReport
{
    public required IReadOnlyList<QuoteLine> Lines;

    // [ExcelsiorTable] can't render a loop-scoped or fully-styled table, so build it directly
    // with Excelsior's WordTableBuilder and return it from an OpenXmlToken. The token's
    // {{ LinesTable }} substitution must sit alone in its paragraph (structural replacement).
    public TokenValue LinesTable =>
        new OpenXmlToken(context =>
        [
            new WordTableBuilder<QuoteLine>(Lines)
                .HeadingParagraphStyle("TBLHeading")
                .BodyParagraphStyle("TBLText")
                .Build(context.MainPart)
        ]);
}
```
<sup><a href='/src/Parchment.Tests/Docx/ExcelsiorTableTests.cs#L156-L175' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcelsiorTableViaOpenXmlToken' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Reference the property with a solo `{{ LinesTable }}` token (it must sit alone in its paragraph — the same structural-replacement rule as any `OpenXmlToken`). The full `WordTableBuilder` surface is available — `headingStyle`/`bodyStyle` callbacks, per-column `CellStyle`, and the `HeadingParagraphStyle`/`BodyParagraphStyle` shown above. Because the property is evaluated per render, the same pattern inside a `{% for %}` loop produces one table per iteration.


### Html and Markdown properties

Mark a `string` property with `[Html]` or `[Markdown]` (any attribute named `HtmlAttribute` / `MarkdownAttribute`, or `[StringSyntax("html")]` / `[StringSyntax("markdown")]`) and the matching `{{ ... }}` substitution is rendered as a structurally-replaced block of Word content instead of raw text. Html runs through the `OpenXmlHtml` converter; markdown runs through the same Markdig pipeline used by the full markdown template flow.

The attributes are detected by name — Parchment does not ship them. Define them in a consuming project (or use `[StringSyntax("html")]` from `System.Diagnostics.CodeAnalysis`):

<!-- snippet: HtmlAttribute -->
<a id='snippet-HtmlAttribute'></a>
```cs
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
sealed class HtmlAttribute : Attribute;
```
<sup><a href='/src/Parchment.Tests/Docx/FormatAttributeTests.cs#L11-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-HtmlAttribute' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Mark the property:

<!-- snippet: HtmlModel -->
<a id='snippet-HtmlModel'></a>
```cs
public class HtmlDoc
{
    public required string Title;

    [Html]
    public required string Body;
}
```
<sup><a href='/src/Parchment.Tests/Docx/FormatAttributeTests.cs#L31-L39' title='Snippet source file'>snippet source</a> | <a href='#snippet-HtmlModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Drop a `{{ Body }}` substitution into the template on its own line:

![Template before render](/src/Parchment.Tests/Scenarios/html-property/input.png)

<!-- snippet: HtmlUsage -->
<a id='snippet-HtmlUsage'></a>
```cs
var templatePath = Path.Combine(ScenarioPath("html-property"), "input.docx");

var store = new TemplateStore();
store.RegisterDocxTemplate<HtmlDoc>("html-doc", templatePath);

var model = new HtmlDoc
{
    Title = "Report",
    Body = "<p>Hello <b>world</b></p><p>Second para</p>"
};

using var stream = new MemoryStream();
await store.Render("html-doc", model, stream);
```
<sup><a href='/src/Parchment.Tests/Docx/FormatAttributeTests.cs#L97-L111' title='Snippet source file'>snippet source</a> | <a href='#snippet-HtmlUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

![Rendered output](/src/Parchment.Tests/Scenarios/html-property/output%23page01.verified.png)

`[Markdown]` is the same shape — mark the property with a `MarkdownAttribute`-named attribute (or `[StringSyntax("markdown")]`) and the string is parsed as markdown at render time:

<!-- snippet: MarkdownModel -->
<a id='snippet-MarkdownModel'></a>
```cs
public class MarkdownDoc
{
    public required string Title;

    [Markdown]
    public required string Body;
}
```
<sup><a href='/src/Parchment.Tests/Docx/FormatAttributeTests.cs#L41-L49' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

![Template before render](/src/Parchment.Tests/Scenarios/markdown-property/input.png)

![Rendered output](/src/Parchment.Tests/Scenarios/markdown-property/output%23page01.verified.png)

As an alternative to defining custom attributes, `[StringSyntax]` from `System.Diagnostics.CodeAnalysis` is equivalent (case-insensitive):

<!-- snippet: StringSyntaxHtmlModel -->
<a id='snippet-StringSyntaxHtmlModel'></a>
```cs
public class StringSyntaxHtmlDoc
{
    public required string Title;

    [StringSyntax("html")]
    public required string Body;
}
```
<sup><a href='/src/Parchment.Tests/Docx/FormatAttributeTests.cs#L51-L59' title='Snippet source file'>snippet source</a> | <a href='#snippet-StringSyntaxHtmlModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

<!-- snippet: StringSyntaxMarkdownModel -->
<a id='snippet-StringSyntaxMarkdownModel'></a>
```cs
public class StringSyntaxMarkdownDoc
{
    public required string Title;

    [StringSyntax("markdown")]
    public required string Body;
}
```
<sup><a href='/src/Parchment.Tests/Docx/FormatAttributeTests.cs#L61-L69' title='Snippet source file'>snippet source</a> | <a href='#snippet-StringSyntaxMarkdownModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Rules:

- The substitution must be a plain member-access expression — filters and arithmetic are rejected because the formatted rendering is selected by attribute rather than Fluid evaluation. Diagnostic `PARCH010`.
- Only `string` / `string?` properties are supported.
- `[Html]` + `[Markdown]` on the same property, or `[Html]` + `[StringSyntax("markdown")]` (and vice versa), is rejected at registration as a mismatch.
- Unlike `[ExcelsiorTable]`, the token does **not** have to sit alone in its paragraph. Surrounding text and sibling tokens are preserved — see [Inline-aware structural replacement](#inline-aware-structural-replacement).

#### Inline-aware structural replacement

When an `[Html]` / `[Markdown]` token shares its paragraph with other text or sibling tokens, Parchment chooses one of three modes based on what the rendered content looks like:

| Token position | Rendered shape | Result |
| --- | --- | --- |
| Solo (covers the whole paragraph) | Anything | Host paragraph is replaced by the rendered elements. The host's `pPr` is lost; rendered paragraphs/tables stand on their own. |
| Non-solo | A single paragraph (typical for inline-only HTML like `<b>x</b>`, or single-line markdown) | The produced paragraph's `pPr` is dropped; its runs are spliced into the host at the token offset. Surrounding text and the host's `pPr` are preserved. |
| Non-solo | Multiple block-level elements, or a non-paragraph block (table) | The host is split at the token offset: text before becomes its own paragraph (cloning host's `pPr`), the rendered blocks slot in between, and text after becomes another paragraph (also cloning host's `pPr`). Empty before/after halves are still emitted. |

Two non-solo block-shaped tokens in the same paragraph are rejected at render time — the splits would overlap and there is no defined way to compose them. Move one of the tokens to its own paragraph.


### Enumerable string properties

Any property assignable to `IEnumerable<string>` (`string[]`, `List<string>`, `IReadOnlyList<string>`, etc.) is auto-rendered as a Word native bullet list when referenced as a solo `{{ Property }}` substitution. **No attribute is required** — detection is purely type-driven, mirroring [Excelsior's Enumerable string properties](https://github.com/Papyrine/Excelsior#enumerable-string-properties) feature.

Mark the property:

<!-- snippet: StringListModel -->
<a id='snippet-StringListModel'></a>
```cs
public class Person
{
    public required string Name;
    public required IEnumerable<string> Tags;
}
```
<sup><a href='/src/Parchment.Tests/Docx/StringListTests.cs#L12-L18' title='Snippet source file'>snippet source</a> | <a href='#snippet-StringListModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Drop a `{{ Tags }}` substitution into the template on its own line:

![Template before render](/src/Parchment.Tests/Scenarios/string-list/input.png)

Register and render normally:

<!-- snippet: StringListUsage -->
<a id='snippet-StringListUsage'></a>
```cs
var templatePath = Path.Combine(ScenarioPath("string-list"), "input.docx");

var store = new TemplateStore();
store.RegisterDocxTemplate<Person>("string-list-scenario", templatePath);

var model = new Person
{
    Name = "Ada Lovelace",
    Tags = ["Author", "Mathematician", "Engineer"]
};

using var stream = new MemoryStream();
await store.Render("string-list-scenario", model, stream);
```
<sup><a href='/src/Parchment.Tests/Docx/StringListTests.cs#L123-L139' title='Snippet source file'>snippet source</a> | <a href='#snippet-StringListUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The rendered output:

![Rendered output](/src/Parchment.Tests/Scenarios/string-list/output%23page01.verified.png)

Behavior:

- The auto path is **opportunistic, not strict**: it only fires when the token sits alone in its paragraph and has no filter chain. Otherwise the substitution falls through to standard Fluid evaluation — so existing `{{ Tags | bullet_list }}` and `{{ Tags | numbered_list }}` filter usage keeps working unchanged.
- A null collection renders as empty (no error). An empty collection renders as no paragraphs.
- Properties marked `[ExcelsiorTable]` keep ownership — the Excelsior path runs first and wins for `[ExcelsiorTable] IEnumerable<string>` (though Excelsior itself rejects `string` as an element type, so this combination is rarely useful in practice).
- Loop-scoped tokens fall through. The detection map is keyed on dotted paths from the root model only, matching the `[ExcelsiorTable]` and `[Html]`/`[Markdown]` limitations. Inside `{% for c in Customers %}{{ c.Tags }}{% endfor %}`, use the explicit `bullet_list` filter.
- Numbered output: opt out of the auto path with `{{ Tags | numbered_list }}`.


## Editable fields (two-way binding)

Mark model members with `[EditableField]` and the rendered docx becomes a fillable, locked form:

- Each marked member renders as a Word **content control** (`w:sdt`) tagged with the member's dotted model path and typed per the member — plain text, checkbox, date picker, dropdown, or a rich-text block for HTML.
- The rest of the document is locked: `w:documentProtection w:edit="readOnly"` is written into the output's settings, with an editable-range exception (`w:permStart`/`w:permEnd`, everyone) punched around each control. Users can fill the fields but not edit anything else; the controls themselves are locked against deletion (`w:lock="sdtLocked"`).
- `ParchmentExtractor` reads the values back out of the returned document and applies them onto a model instance — two-way binding for the editable subset. The document is a lossy projection of the model (loops expanded, formatting applied), so it is the *editable subset* that round-trips, merged onto a caller-supplied model.

Model:

<!-- snippet: EditableFieldsModel -->
<a id='snippet-EditableFieldsModel'></a>
```cs
public class OrderForm
{
    public required string Number;

    [EditableField]
    public required string PurchaseOrder { get; set; }

    [EditableField]
    public bool Approved { get; set; }

    [EditableField]
    public OrderStatus Status { get; set; }

    [EditableField(MultiLine = true)]
    public string? Notes { get; set; }
}

public enum OrderStatus
{
    Draft,
    Submitted,
    Accepted
}
```
<sup><a href='/src/Parchment.Tests/Docx/EditableFieldTests.cs#L13-L37' title='Snippet source file'>snippet source</a> | <a href='#snippet-EditableFieldsModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The template ([input.docx](/src/Parchment.Tests/Scenarios/editable-fields/input.docx)):

![Template before render](/src/Parchment.Tests/Scenarios/editable-fields/input.png)

Register and render normally:

<!-- snippet: EditableFieldsUsage -->
<a id='snippet-EditableFieldsUsage'></a>
```cs
var templatePath = Path.Combine(ScenarioPath("editable-fields"), "input.docx");

var store = new TemplateStore();
store.RegisterDocxTemplate<OrderForm>("order-form", templatePath);

var model = new OrderForm
{
    Number = "ORD-2026-042",
    PurchaseOrder = "PO-77041",
    Approved = true,
    Status = OrderStatus.Submitted,
    Notes = null
};

using var stream = new MemoryStream();
await store.Render("order-form", model, stream);
```
<sup><a href='/src/Parchment.Tests/Docx/EditableFieldTests.cs#L42-L59' title='Snippet source file'>snippet source</a> | <a href='#snippet-EditableFieldsUsage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The rendered output — editable fields as content controls, everything else read-only:

![Rendered output](/src/Parchment.Tests/Scenarios/editable-fields/output.verified.png)

After the user fills in the form in Word, extract the edited values back onto a model:

<!-- snippet: EditableFieldsExtract -->
<a id='snippet-EditableFieldsExtract'></a>
```cs
var result = ParchmentExtractor.Extract<OrderForm>(stream);

result.ApplyTo(model);
```
<sup><a href='/src/Parchment.Tests/Docx/EditableFieldTests.cs#L102-L106' title='Snippet source file'>snippet source</a> | <a href='#snippet-EditableFieldsExtract' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`ExtractResult.Fields` reports a per-field outcome — `Extracted`, `Empty` (placeholder showing), `Missing` (control deleted, or the document wasn't produced from a template bound to this model), `ParseFailed` (with the raw text for diagnostics), or `Duplicate` — and `AllExtracted` is true when every field read cleanly. `ApplyTo` assigns `Extracted` values, applies `Empty` as null to nullable members only, and validates that every intermediate object on nested paths is non-null *before* mutating anything (throwing `ParchmentExtractionException` otherwise). Extraction needs no `TemplateStore` — controls are matched by tag, so a service that only receives filled-in documents never has to register the template.

### Control kinds

| Member type | Control | Round-trip value |
| --- | --- | --- |
| `string` | plain text (`MultiLine = true` allows line breaks) | the text |
| `bool` | checkbox | `w14:checked` — no text parsing |
| `DateOnly` / `DateTime` | date picker (`DateFormat` option, default `yyyy-MM-dd`) | canonical `w:fullDate`, maintained by Word's picker |
| `DateTimeOffset` | plain text (round-trippable ISO, default `yyyy-MM-ddTHH:mm:sszzz`) | the ISO text, parsed offset-preserving |
| `TimeOnly` | plain text (default `HH:mm:ss`) | the text |
| enums | dropdown, one item per member | `w:listItem` value |
| numeric types | plain text | text parsed with the extraction culture |
| `string` marked `[Html]` | rich-text block control (block `w:sdt`) holding editable formatted content | HTML serialized back from the block |

Nullable variants are supported except `bool?` — a checkbox cannot represent null. Null values render as the control's grey placeholder and extract as `Empty`. Members must have a public non-init setter so extraction can write back.

**Why the temporal split.** `DateOnly` and `DateTime` use Word's native date picker, whose `w:fullDate` is a canonical value the picker maintains — so those never depend on parsing display text. `DateTimeOffset` and `TimeOnly` get **plain-text** controls instead: Word has no offset-aware or time-only picker, and `w:fullDate` is a bare `DateTime` that cannot carry an offset. Their run text is therefore the source of truth, rendered in a round-trippable ISO format and parsed back with the offset intact. Two consequences worth knowing:

- A `DateTime` comes back with `DateTimeKind.Unspecified` (`w:fullDate` carries no zone) — re-stamp the `Kind` where a specific one is required. Its time-of-day survives in `w:fullDate` even though the default `yyyy-MM-dd` format hides it; a time-bearing `DateFormat` (e.g. `"yyyy-MM-dd HH:mm"`) makes the time visible and editable as text.
- The `DateTimeOffset` / `TimeOnly` defaults keep seconds but not sub-second precision; a `DateFormat` such as `"o"` preserves it. Because these parse display text, the extraction culture must match the render culture (as with numerics).

### Rich text (HTML)

A `string` member marked both `[EditableField]` and `[Html]` renders as an **editable rich-content block**. The HTML value becomes formatted Word content — paragraphs, bold / italic / underline / strikethrough, superscript / subscript, bullet and numbered lists, links, and `h1`–`h6` headings — inside a block-level content control wrapped in an editable range, so the formatted content stays editable while the rest of the document remains locked. Extraction serializes the (possibly Word-edited) block back to an HTML string, so the field round-trips as HTML rather than flattened text:

```cs
public class Article
{
    [Html]
    [EditableField]
    public required string Body { get; set; }
}
```

The token must sit alone in its paragraph (the control is block-level, matching the read-only `[Html]` rule). The round-trip covers the subset a rich-text editor emits; content outside that subset degrades to its text. `[Markdown]` combined with `[EditableField]` is rejected (`PARCH015`) — editable round-trip is HTML-only, since extraction has no OpenXML-to-Markdown serializer.

When a model has an `[EditableField] [Html]` member, registration seeds a bullet and a numbered list definition into the template (via `WordNumbering.EnsureListDefinitions`). Word can apply a list only against a definition that already exists — creating one writes to the document-level `word/numbering.xml`, which read-only protection locks — so without the seed the bullet and numbering buttons stay greyed out even inside the editable block. Seeding a definition Word can reference makes the buttons live, so a user can add a list to a field that started with none. The definitions are unused clutter if no list is ever added, which is harmless; the seed is skipped entirely when protection is off (`ProtectionMode.None`) or no member is `[Html]`.

### Protection is cooperative, not security

Word enforces `w:documentProtection` in its UI only. No password is set (a password would add no real protection while forcing non-deterministic salt generation into the output), so any user can lift the protection via Review → Restrict Editing, and any program can edit the underlying XML regardless. Editable fields are the right tool for *guiding* users to fill in specific values — not for tamper-proofing.

Protection is applied at registration whenever the model declares at least one `[EditableField]` member. Opt out per template:

```csharp
store.RegisterDocxTemplate<OrderForm>("order-form", path, ProtectionMode.None);
```

or via the source-generator attribute:

```csharp
[ParchmentModel("Templates/order-form.docx", Protection = ProtectionMode.None)]
```

With `ProtectionMode.None` the fields still render as tagged controls (and still extract); the rest of the document stays editable too.

### Plain-text fields accept character formatting

A plain `[EditableField]` renders as a Word plain-text content control (`<w:text/>`). That element constrains the control's *content model* — a single run of text, no images, no nested controls, no extra paragraphs — but it says nothing about run properties, and Word deliberately lets a user bold or italicise the control's contents. Nothing on `w:sdt` forbids it: `w:lock="sdtLocked"` only stops the control being deleted.

The formatting is cosmetic. Extraction reads the control's `w:t` text (turning `w:br` into a newline) and ignores `w:rPr` entirely, so a bolded field extracts exactly as an unbolded one and the emphasis never reaches the model. This applies to plain fields only — for an `[EditableField] [Html]` field, formatting *is* the value, and bold round-trips as `<strong>`.

The one way to make Word refuse the keystroke is `w:documentProtection w:edit="forms"`, which disallows formatting inside the protected region. Parchment writes `readOnly` instead, because forms protection also blocks selecting and copying the locked text.

### Rich content cannot be whitelisted in Word

There is no way to tell Word "allow bold and italic in this control, but not tables". The file format has no per-control formatting whitelist, and the two protection dials that come close are both the wrong shape:

- **`w:edit="forms"` is all-or-nothing.** Users may fill in fields and nothing else: no bold, no italic, no table insertion. There is no way to permit character formatting while denying block content.
- **`w:formatting="1"` is style-granular, not element-granular.** It means *only allow formatting with unlocked styles* — the unit of permission is a style (a style carrying `<w:locked/>` cannot be applied), and switching it on blocks direct formatting outright. Bold and italic are direct formatting rather than styles, and a table is not a style at all, so neither can be expressed. It is orthogonal to `w:edit`.

`w:sdt` itself offers only a fixed menu of discrete control types — plain text, rich text, combo box, dropdown, date, checkbox, picture. "Rich text" means any content: paragraphs, images, tables. There is no inline-emphasis-only variant.

So the boundary is enforced on extraction, not in Word. Whatever a user pastes into an `[EditableField] [Html]` control, only the supported subset survives the round-trip: a table degrades to its cell paragraphs (no text is dropped, but it never returns as `<table>`), and other unsupported content degrades to its text. Tighten the rule further — stripping `<strong>`, say — in the serializer, not in the protection settings.

### Culture

Checkbox, date, and dropdown values round-trip through canonical control state and never depend on display text. String content is taken verbatim. Only free-text *numerics* are culture-sensitive: they render via the Fluid template culture (invariant by default) and parse back with `Extract`'s `culture` parameter — which defaults to `CultureInfo.InvariantCulture` to match. If renders are configured for another culture, pass the same culture to `Extract`.

### Rules and limitations

- **Docx flow only.** Markdown templates substitute into text before parsing, so substitution sites don't exist in the output; `[EditableField]` is ignored there — same rule as `[ExcelsiorTable]` and `[Html]`/`[Markdown]`.
- **Document body only.** A token referencing an editable member in a header or footer renders as a plain read-only substitution — deliberate, so an editable body field can be mirrored read-only in a footer. Extraction reads the single body control.
- **Loops render read-only.** An editable token inside `{% for %}` would stamp one control per iteration and produce duplicate tags, so loop bodies render editable members as plain text (`PARCH018` warns at compile time).
- **One token per member** in the body (`PARCH017`) — the dotted path is the control's tag and must be unique for extraction. Plain member access only; filters are rejected (`PARCH016`).
- **No structural mixing.** An editable token cannot share a paragraph with an `[ExcelsiorTable]` / `[Html]` / `[Markdown]` token — structural replacement clones paragraph halves and would corrupt the editable range. Rejected at registration.
- `[EditableField]` combined with `[ExcelsiorTable]` or `[Markdown]` on the same member is rejected (`PARCH015`); `[Html]` (or `[StringSyntax("html")]`) instead selects the rich-text kind (see [Rich text (HTML)](#rich-text-html)). Unsupported member types are rejected (`PARCH013`); a missing setter is rejected (`PARCH014`). All fail fast at runtime registration.
- **Static members are skipped**, matching the other per-template maps — see [Model binding limitations](#model-binding-limitations).

### Editable collections

Mark a collection member — a `List<T>`, `T[]`, or `IList<T>` of a POCO element type — with `[EditableField]`, and a `{% for %}` loop over it renders as a **Word repeating section** (`w15:repeatingSection`). Reviewers add, remove, and reorder rows with Word's native repeating-section controls; each row's fields are ordinary editable controls, and `ParchmentExtractor` rebuilds the list from whatever rows the document ends up with.

```cs
public class Milestone
{
    [EditableField]
    public string Name { get; set; } = "";

    [EditableField]
    public DateOnly Due { get; set; }
}

public class Roadmap
{
    [EditableField]
    public List<Milestone> Milestones { get; set; } = [];
}
```

The template is a plain loop; the loop-variable tokens become the per-row controls:

```liquid
{% for m in Milestones %}
{{ m.Name }} — due {{ m.Due }}
{% endfor %}
```

After the form comes back, extract onto a model instance — the rebuilt list replaces the collection member:

```cs
var result = ParchmentExtractor.Extract<Roadmap>(stream);
result.ApplyTo(model); // model.Milestones is the edited list — added rows appear, removed rows are gone
```

- **Replace-all round-trip.** The document's rows are the new state. A caller that preserves entity identity (matching rebuilt items onto tracked records by id) performs that mapping itself — the library does not infer identity.
- **An empty collection renders one blank row** so Word has a template to clone; an all-blank row extracts as nothing, so clearing a row deletes it.
- **The element type must be reconstructable**: a public parameterless constructor and settable `[EditableField]` members, plus a settable collection member. One nesting level — an element type cannot itself hold an editable collection.
- **Scalar collections** (`string[]`, `List<int>`) are modelled by wrapping the value in a single-field element type, so every repeating section carries named, item-relative control tags.
- **Protection stays `readOnly`.** The repeating section sits inside a `permStart`/`permEnd` range, so add/remove works under the same lock-down as every other control. `readOnly` (rather than fill-in-forms protection) is deliberate: it keeps the locked text selectable and copyable, which forms protection blocks.


## Markdown template

A markdown template is a `.md` file containing the full body of the document plus liquid tokens for substitution, looping, and conditional content. The template below combines headings, emphasis, a pipe table driven by a loop, an ordered list driven by a loop, and a blockquote chosen by an `{% if %}`:


### Sample

<!-- snippet: MarkdownTemplate -->
<a id='snippet-MarkdownTemplate'></a>
```handlebars
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
```
<sup><a href='/src/Parchment.Tests/UsageTests.cs#L60-L88' title='Snippet source file'>snippet source</a> | <a href='#snippet-MarkdownTemplate' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The model the template binds against:

<!-- snippet: ReportModel -->
<a id='snippet-ReportModel'></a>
```cs
public class ReportContext
{
    public required Report Report;
}

public class Report
{
    public required string Title;
    public required string Author;
    public required Date Date;
    public required string Summary;
    public required IReadOnlyList<Finding> Findings;
    public required IReadOnlyList<ActionItem> Actions;
    public required bool HasRisks;
}

public class Finding
{
    public required string Area;
    public required string Status;
    public required string Owner;
}

public class ActionItem
{
    public required string Title;
    public required string Detail;
}
```
<sup><a href='/src/ParchmentModel/Report.cs#L3-L32' title='Snippet source file'>snippet source</a> | <a href='#snippet-ReportModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render it like any other template:

<!-- snippet: Markdown -->
<a id='snippet-Markdown'></a>
```cs
using var brandDocx = DocxTemplateBuilder.Build();
var reportModel = SampleData.Report();

var store = new TemplateStore();
store.RegisterMarkdownTemplate<ReportContext>(
    "report",
    markdownSource,
    styleSource: brandDocx);

using var stream = new MemoryStream();
await store.Render("report", reportModel, stream);
```
<sup><a href='/src/Parchment.Tests/UsageTests.cs#L91-L105' title='Snippet source file'>snippet source</a> | <a href='#snippet-Markdown' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The rendered docx (page 1):

![Markdown template output](/src/Parchment.Tests/UsageTests.Markdown%23page01.verified.png)

The optional `styleSource` is a docx whose styles, headers, footers, theme, and section properties (page size, margins, header/footer references) are inherited by the output. If omitted, a built-in blank template is used.


### Supported Markdig extensions

#### [Emphasis extras](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/EmphasisExtraSpecs.md)

Extended emphasis syntax beyond standard bold/italic:

```markdown
~~strikethrough~~
~subscript~
^superscript^
++underline++
==highlight==
```

Standard `*italic*`, `**bold**`, and `_italic_` work as usual.

#### [Pipe tables](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/PipeTableSpecs.md)

```markdown
| A | B |
|---|---|
| 1 | 2 |
| 3 | 4 |
```

Header cells are bold and centered. Rendered output:

![Pipe table output](/src/Parchment.Tests/Markdown/Renderers/TableRendererTests.PipeTableEmitsTableWithGridRowsAndHeaderFormatting%23page01.verified.png)

Column alignment markers in the header separator row (`:---`, `:---:`, `---:`) are honored for both header and body cells:

```markdown
| Left | Center | Right | Default |
|:-----|:------:|------:|---------|
| a    | b      | c     | d       |
| e    | f      | g     | h       |
```

Columns with no alignment marker fall back to the defaults (header centered, body left). Rendered output:

![Aligned pipe table output](/src/Parchment.Tests/Markdown/Renderers/TableRendererTests.PipeTableHonorsColumnAlignment.verified.png)

The number of dashes in each separator cell determines the column's relative width. The table uses fixed layout and each `gridCol` / cell gets an explicit width proportional to its dash count:

```markdown
| Name | Description           | Count |
|------|--------------------|------|
| A    | Short                 | 1     |
| BB   | A longer description  | 22    |
```

Dash counts 6 / 20 / 6 yield widths in the same ratio (18.75% / 62.5% / 18.75%). Rendered output:

![Pipe table column widths](/src/Parchment.Tests/Markdown/Renderers/TableRendererTests.PipeTableHonorsColumnWidthsFromDashCounts.verified.png)

Behaviour notes:

- If the `|` characters in the header, separator, and body rows all line up at the same column positions, dash counts are treated as readability padding and ignored — the table sizes naturally. Custom widths are inferred only when the separator dashes intentionally break that alignment. The two snippets below illustrate the distinction:

  ```markdown
  | Left | Center | Right | Default |    ← pipes align across all rows ⇒ uniform widths
  |:-----|:------:|------:|---------|
  | a    | b      | c     | d       |
  ```

  ```markdown
  | Name | Description           | Count |    ← separator pipes shift ⇒ widths apply
  |------|--------------------|------|
  | A    | Short                 | 1     |
  ```

- When all column widths are equal (uniform separators like `|---|---|`), no explicit widths are emitted — Word's default auto-distribution produces the same layout, so the docx stays minimal.
- Tables nested in blockquotes or list items are auto-sized to fit their indented container, so the explicit dxa column widths are skipped regardless of dash counts.

#### [Grid tables](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/GridTableSpecs.md)

Grid tables use `+---+` borders and `+===+` to separate the header row:

```markdown
+---+---+
| A | B |
+===+===+
| 1 | 2 |
+---+---+
| 3 | 4 |
+---+---+
```

Rendered output:

![Grid table output](/src/Parchment.Tests/Markdown/Renderers/TableRendererTests.GridTableEmitsTableWithCorrectStructure%23page01.verified.png)

#### [Auto links](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/AutoLinks.md)

Bare URLs and email addresses are automatically converted to hyperlinks with the `Hyperlink` character style:

```markdown
Visit https://example.com or email user@example.com
```

#### [List extras](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/ListExtraSpecs.md)

Alpha and roman numeral list markers beyond standard `1.` numbering:

```markdown
a. lower alpha
b. items

A. upper alpha
B. items

i. lower roman
ii. items

I. upper roman
II. items
```

Each format produces the corresponding Word numbering definition. Rendered output (lower alpha):

![Lower alpha list output](/src/Parchment.Tests/Markdown/Renderers/ListBlockRendererTests.LowerAlphaListUsesLowerLetterFormat%23page01.verified.png)

#### [Smarty pants](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/SmartyPantsSpecs.md)

ASCII quotes and dashes are replaced with typographic equivalents:

| Input | Output |
|---|---|
| `'text'` | \u2018text\u2019 (curly single quotes) |
| `"text"` | \u201Ctext\u201D (curly double quotes) |
| `--` | \u2013 (en-dash) |
| `---` | \u2014 (em-dash) |
| `...` | \u2026 (ellipsis) |

#### [Generic attributes](https://github.com/xoofx/markdig/blob/main/src/Markdig.Tests/Specs/GenericAttributesSpecs.md)

Attach a Word style to a heading or paragraph with `{.StyleName}` syntax. The first class attribute is used as the paragraph's `ParagraphStyleId`:

```markdown
## Section heading {.MyCustomHeading}

Some intro paragraph. {.IntroBlock}
```


### HTML comments are stripped

HTML comment blocks (`<!-- ... -->`) are dropped during rendering rather than passed through as empty paragraphs. This allows embedding snippet markers, authoring notes, or TODOs in template sources without bleeding visible whitespace into the output docx:

```markdown
# {{ Title }}

<!-- TODO: add executive summary -->
Body text follows the heading.
```

Only standalone comment *blocks* are removed; inline HTML, scripts, styles, and any other HTML constructs render normally via [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml).


## Images

Parchment has no dedicated image binding or attribute. Images flow through one of three paths depending on where the reference lives.


### Static images embedded in the template

Pictures placed directly in the `.docx` template (Word's Insert &rarr; Picture, or pasted in) pass through to the output verbatim. Token scanning only mutates paragraphs that contain liquid tokens; image parts, drawings, and shapes elsewhere in the document are cloned byte-for-byte. Drop a logo into the header in Word, render, the logo comes out — no model wiring needed.


### `<img>` inside HTML and markdown `![alt](url)`

`<img>` tags (in an `[Html]` property, an HTML block inside a markdown template, or inline HTML inside a `[Markdown]` property) and markdown `![alt](url)` syntax both delegate to [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml)'s `ImageResolver`. The same resolution table applies to both:

| `src` / `url` value | Result |
| --- | --- |
| `data:image/...;base64,...` (data URI) | Bytes decoded, embedded as `<w:drawing>` |
| `file:///...` URI | File read from disk, embedded |
| Absolute or CWD-relative file path (e.g. `C:\images\logo.png`, `images/logo.png`) | File read from disk, embedded |
| `http://...` / `https://...` | Fetched synchronously, embedded |
| Path that fails the active `ImagePolicy` (or fails to read) | Falls back to rendering the alt text |

By default, `TemplateStore` exposes `LocalImages = ImagePolicy.AllowAll()` and `WebImages = ImagePolicy.AllowAll()` — Parchment renders developer-bound model content, so locking either down by default would silently break `<img src="C:\...">` and `![](path)` references. If a model carries images from less-trusted sources, override the policies on the store:

```cs
var store = new TemplateStore
{
    LocalImages = ImagePolicy.SafeDirectories("C:/assets/branding"),
    WebImages = ImagePolicy.Deny()
};
```

The `ImagePolicy` type is OpenXmlHtml's — `Deny`, `AllowAll`, `SafeDomains(...)`, `SafeDirectories(...)`, `Filter(predicate)`. See [OpenXmlHtml's image-policy docs](https://github.com/Papyrine/OpenXmlHtml#image-policy) for the full surface.

Determinism holds as long as the bytes at the resolved source don't change between renders. Enabling `WebImages` couples render output to network state — if reproducibility matters, materialize web URLs into the model (or use `ImagePolicy.Deny()` for `WebImages`) before rendering.


### Custom `OpenXmlToken` (programmatic embedding)

For full control — explicit sizing, anchored positioning, charts, anything outside markdown or HTML — return an `OpenXmlToken` from a `TokenValue`-typed model property. The render delegate receives an `IOpenXmlContext` whose `AddImagePart(byte[] bytes, string contentType)` adds an `ImagePart` to the document and returns its relationship ID. Build the `<w:drawing>` directly (inline or anchor element, EMU-sized extents, blip ref pointing at the rel-id) and yield it as one of the produced elements.

Use this path when image bytes live in memory (rendered chart, database blob, dynamically generated PNG) and round-tripping through base64 or a temp file would be wasteful.

Model:

<!-- snippet: ImageTokenModel -->
<a id='snippet-ImageTokenModel'></a>
```cs
public class BrandKit
{
    public required string Title;
    public required TokenValue Logo;
}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L291-L299' title='Snippet source file'>snippet source</a> | <a href='#snippet-ImageTokenModel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Word's `Drawing` element conflicts with the `DocumentFormat.OpenXml.Drawing` namespace, so namespace aliases keep the construction code readable:

<!-- snippet: ImageTokenAliases -->
<a id='snippet-ImageTokenAliases'></a>
```cs
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L1-L5' title='Snippet source file'>snippet source</a> | <a href='#snippet-ImageTokenAliases' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Template content:

<!-- snippet: ImageTokenContent -->
<a id='snippet-ImageTokenContent'></a>
```cs
# {{ Title }}

{{ Logo }}
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L307-L311' title='Snippet source file'>snippet source</a> | <a href='#snippet-ImageTokenContent' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Render — `context.AddImagePart` returns a relationship ID; the rest is standard OpenXML drawing plumbing (extent in EMUs, blip referencing the rel-id, preset shape geometry):

<!-- snippet: ImageTokenRender -->
<a id='snippet-ImageTokenRender'></a>
```cs
var store = new TemplateStore();
store.RegisterDocxTemplate<BrandKit>("image-token", template);
await store.Render(
    "image-token",
    new BrandKit
    {
        Title = "Brand kit",
        Logo = new OpenXmlToken(context =>
        {
            var relId = context.AddImagePart(imageBytes, "image/png");

            // Word measures images in EMUs (English Metric Units): 914400 per inch.
            const long widthEmu = 914400L;
            const long heightEmu = 914400L;

            var inline = new DW.Inline(
                new DW.Extent
                {
                    Cx = widthEmu,
                    Cy = heightEmu
                },
                new DW.DocProperties
                {
                    Id = 1U,
                    Name = "Logo"
                },
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties
                                {
                                    Id = 0U,
                                    Name = "logo.png"
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip
                                {
                                    Embed = relId
                                },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset
                                    {
                                        X = 0L,
                                        Y = 0L
                                    },
                                    new A.Extents
                                    {
                                        Cx = widthEmu,
                                        Cy = heightEmu
                                    }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }));

            return [new Paragraph(new Run(new Drawing(inline)))];
        })
    },
    stream);
```
<sup><a href='/src/Parchment.Tests/Docx/TokenOverrideTests.cs#L318-L387' title='Snippet source file'>snippet source</a> | <a href='#snippet-ImageTokenRender' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


## Registration-time validation

Whether registering by hand (`RegisterDocxTemplate<T>(...)`) or through the source generator's `RegisterWith(store)` helper, the template is fully validated against `T` at registration — before any render runs. Missing members, block tags targeting non-enumerable properties, or `[ExcelsiorTable]` tokens that break the solo-in-paragraph / plain-member-access rules throw `ParchmentRegistrationException` immediately. Register templates at app startup and any binding mismatch surfaces there, not on the first render.


## Source generator (recommended)

The source generator is the recommended way to register templates. Decorate the model class itself with `[ParchmentModel]` and Parchment's source generator (a) validates the template tokens against the model at compile time, and (b) emits a `RegisterWith` helper that pre-compiles all the member accessors needed at registration time — Fluid accessors per reachable type, plus the per-template maps for `[ExcelsiorTable]`, `[Html]` / `[Markdown]`, and `IEnumerable<string>` properties.

**No runtime reflection at registration.** Both the model graph walk in `SharedFluid.RegisterTypeGraph` and the dotted-path walks in `ExcelsiorTableMap` / `FormatMap` / `StringListMap` short-circuit when the source generator has already populated their caches. This makes Parchment safe for trimming and NativeAOT scenarios, and matches the modern .NET pattern set by `System.Text.Json`'s `JsonSerializerContext`, EF Core's compiled models, and the regex source generator.

The model must be `partial` — the generator emits a `RegisterWith` helper onto the same class. Both docx and markdown templates are supported — the generator branches on the path's extension (`.docx` → docx flow, `.md` → markdown flow):

```csharp
[ParchmentModel("Templates/invoice.docx")]
public partial class Invoice
{
    public string Number { get; set; } = "";
    public Customer Customer { get; set; } = new();
    // ...
}

[ParchmentModel("Templates/report.md")]
public partial class Report
{
    public string Title { get; set; } = "";
    // ...
}
```

The attribute is applied directly to the binding model — there is no separate marker / "template" class. Models almost always need Parchment-aware code on them anyway (`[Html]` / `[Markdown]` / `[ExcelsiorTable]` annotations, helper properties shaping values for binding), so the `partial` + Parchment-dependency tax is already paid. See `CLAUDE.md` → "Design decisions" for the full rationale.

The runtime `TemplateStore.RegisterDocxTemplate<T>(name, path)` / `RegisterMarkdownTemplate<T>(name, markdown)` overloads still exist as an **escape hatch** — use them when the model can't be made `partial` (third-party / generated types), when the model isn't known at compile time, or when one model needs to bind multiple templates. They are reflection-based and not trimming-friendly; the SG path is preferred whenever the model is yours to annotate.

In both cases the generator also emits a `RegisterWith(store)` helper so registration is one line at runtime:

```csharp
var store = new TemplateStore();
Invoice.RegisterWith(store);
Report.RegisterWith(store, styleSource: File.OpenRead("brand.docx"));
```

The markdown helper has an extra optional `styleSource` parameter that mirrors `RegisterMarkdownTemplate<T>` — pass a brand docx whose page setup, headers/footers, and styles should be inherited by the rendered output.

Add each template to `<AdditionalFiles>` so the generator can find it:

```xml
<ItemGroup>
  <AdditionalFiles Include="Templates\invoice.docx" />
  <AdditionalFiles Include="Templates\report.md" />
</ItemGroup>
```

The generator emits the following diagnostics. Unless noted, each applies to both flows.


### `PARCH001` — unknown model member

A `{{ }}` token references a property that doesn't exist on the model type.

```
// Model
public class Letter
{
    public Customer Customer { get; set; }
}

public class Customer
{
    public string Name { get; set; }
}

// Template paragraph — "Missing" does not exist on Customer
{{ Customer.Missing }}
```


### `PARCH002` — loop source is not enumerable

A `{% for %}` tag targets a property that doesn't implement `IEnumerable<T>`. Note that `string` is explicitly rejected even though it implements `IEnumerable<char>`.

```
// Model — Customer is not a collection
public class Letter
{
    public Customer Customer { get; set; }
}

// Template paragraphs
{% for item in Customer %}
...
{% endfor %}
```


### `PARCH003` — unsupported block tag

Only `for`/`endfor`/`if`/`elsif`/`else`/`endif` are supported as block tags.

```
// Template paragraph — "foobar" is not a recognised tag
{% foobar %}
```


### `PARCH004` — template file not in `<AdditionalFiles>`

The path in `[ParchmentModel("...")]` wasn't found among the project's `<AdditionalFiles>`. Add the docx to the csproj:

```xml
<ItemGroup>
  <AdditionalFiles Include="Templates\invoice.docx" />
</ItemGroup>
```


### `PARCH005` — block tag shares a paragraph

**Docx only.** Block tags (`{% for %}`, `{% if %}`, etc.) must sit alone in their own paragraph. Mixing a block tag with other text on the same line is rejected. Markdown templates are exempt — Fluid parses the whole file as one template, so inline block tags like `Hello {% if x %}World{% endif %}` are legal.

```
// Template paragraphs — "prefix" is on the same line as the for tag
prefix {% for line in Lines %}
{{ line.Description }}
{% endfor %}
```


### `PARCH006` — template file unreadable

The template at the path exists in `<AdditionalFiles>` but couldn't be opened. For docx, typically a corrupt or truncated file. For markdown, this also fires when Fluid fails to parse the file (e.g. an unclosed `{% for %}` block) — the diagnostic message includes the parser error.


### `PARCH007` — `[ExcelsiorTable]` token not alone in paragraph

**Docx only.** `[ExcelsiorTable]`-driven structural replacement is wired through the docx flow only; markdown templates ignore the attribute. An `[ExcelsiorTable]` substitution replaces the entire host paragraph with a Word table. If the paragraph contains other text, that text would be discarded. The token must be the only content in its paragraph.

```
// Model
public class Invoice
{
    [ExcelsiorTable]
    public List<Line> Lines { get; set; }
}

// Template paragraph — "Prefix" shares the paragraph with {{ Lines }}
Prefix {{ Lines }}
```


### `PARCH008` — `[ExcelsiorTable]` token with filters or complex expression

**Docx only.** The Excelsior render path walks the CLR model directly and bypasses Fluid evaluation, so filters and complex expressions would be silently ignored. Only plain member-access (`{{ Lines }}` or `{{ Customer.Lines }}`) is allowed.

```
// Template paragraph — the | reverse filter would be silently dropped
{{ Lines | reverse }}
```


### `PARCH009` — retired

Previously emitted when an `[Html]` / `[Markdown]` token shared its paragraph with other content. The runtime now splices inline content in place and splits the host paragraph for block-level content, so non-solo tokens are valid. The id is intentionally not reused.


### `PARCH010` — `[Html]` / `[Markdown]` token with filters or complex expression

**Docx only.** Formatted rendering is selected by the property attribute rather than via Fluid, so filter chains are not applied. Use plain member access (`{{ Body }}` or `{{ Customer.Bio }}`). The `[Html]` / `[Markdown]` attributes are ignored in the markdown flow — markdown templates have no concept of structural property substitution.


### `PARCH011` — enclosing type of `[ParchmentModel]` target must be partial

The decorated model is nested inside another type that is not declared `partial`. The generator emits the registration helper as `partial class { ... }` and every enclosing type in the chain must be partial too, otherwise the C# compiler rejects the declaration as conflicting with the user's existing one (CS0260). Make every enclosing type partial:

```csharp
public partial class Outer  // <-- partial
{
    [ParchmentModel("template.docx")]
    public partial class Letter
    {
        public Customer Customer { get; set; } = new();
    }
}
```


### `PARCH012` — template missing 'Remove personal information on save' setting

**Docx only. Warning, not error.** The template's `word/settings.xml` does not include the `<w:removePersonalInformation/>` element, which corresponds to Word's "Remove personal information from file properties on save" privacy option. Renders inherit the template's `settings.xml`, so any author / lastModifiedBy / revision metadata on the template propagates to every generated docx. To silence the warning, open the template in Word, go to File → Options → Trust Center → Trust Center Settings → Privacy Options, tick "Remove personal information from file properties on save", and re-save.


### `PARCH013` — `[EditableField]` member has an unsupported type

**Docx only.** Supported types: `string`, `bool`, `DateOnly`, `DateTime`, `DateTimeOffset`, `TimeOnly`, enums, and the numeric primitives / `decimal` — plus nullable variants, except `bool?` (a checkbox cannot represent null).

```csharp
public partial class Order
{
    [EditableField]
    public List<string> Items { get; set; }   // collections are not editable fields
}
```


### `PARCH014` — `[EditableField]` member has no usable setter

**Docx only.** Extraction writes values back onto the model, so the member needs a public non-init setter (or a non-readonly field). `{ get; }`, `{ get; init; }`, expression-bodied getters, and readonly fields are rejected.


### `PARCH015` — `[EditableField]` combined with a conflicting attribute

**Docx only.** `[Html]` (or `[StringSyntax("html")]`) is supported and selects the [rich-text kind](#rich-text-html). Combining `[EditableField]` with `[ExcelsiorTable]`, `[Markdown]`, or `[StringSyntax("markdown")]` on the same member is rejected — editable round-trip is HTML-only (extraction has no OpenXML-to-Markdown serializer).


### `PARCH016` — `[EditableField]` token with filters or complex expression

**Docx only.** The editable render path is selected by attribute rather than via Fluid, so filter chains would be silently ignored. Use plain member access (`{{ PurchaseOrder }}` or `{{ Customer.Notes }}`).


### `PARCH017` — `[EditableField]` member referenced more than once in the body

**Docx only.** The dotted path is the content control's tag, and extraction requires tags to be unique. Reference each editable member once in the document body. (A header or footer occurrence is fine — it renders as a plain read-only mirror and doesn't count.)


### `PARCH018` — `[EditableField]` token inside a loop

**Docx only. Warning, not error.** Loop iterations would stamp one control per iteration and produce duplicate tags, so editable tokens inside `{% for %}` bodies render as plain read-only text instead of editable fields.

```
{% for line in Lines %}
{{ PurchaseOrder }}   ← renders read-only inside the loop
{% endfor %}
```


## Model binding limitations

Parchment binds tokens by reflecting on the model type. Public properties and public fields — both **instance** and **static** — are bindable at every depth, including nested traversal like `{{ Customer.Address.City }}` whether each hop is a property or a field. The source generator's `ShapeBuilder` mirrors the same rules. The following kinds of members are **not** bound — a token referencing them fails registration (`ParchmentRegistrationException`) or compile-time validation (`PARCH001`).

**Static-member caveat**: static members participate in Fluid substitution (`{{ Logo }}` against `public static string Logo`) but **do not** participate in the per-template maps. `[ExcelsiorTable]` on a static collection, `[Html]` / `[Markdown]` on a static string, `[EditableField]` on a static member, and auto-bullet-list dispatch on a static `IEnumerable<string>` are silently treated as no-ops — the runtime map walkers and the SG dotted-path walker both skip static members. Mark the member instance, or wrap it in an instance computed property, when attribute-driven dispatch is required.

### Interfaces as the binding model

`RegisterDocxTemplate<IFoo>(...)` / `RegisterMarkdownTemplate<IFoo>(...)` is rejected at registration time with a `ParchmentRegistrationException` ("Model type 'IFoo' is an interface. Parchment binds against a concrete type's properties via reflection — register against a class, record, or struct instead."). Reflection on an interface returns only members declared directly on that interface, missing inherited base-interface members. The source-generator path is blocked at the language level — `[ParchmentModel]` is declared with `AttributeUsage(AttributeTargets.Class)`, so applying it to an interface is a CS0592 compile error. Abstract classes **are** supported and work as a polymorphic binding surface: register against the abstract type and pass any concrete subclass to `Render`.

### Default implementations on interfaces

A property declared with a body on an interface — e.g. `interface IDocument { string Title { get; } string Header => $"=== {Title} ==="; }` — is **not** bound on classes that implement the interface unless the class re-declares it as a regular instance property. `typeof(Foo).GetProperties(...)` does not surface interface-default members; they're only reachable via the interface type. The SG-side walker visits `BaseType` but not implemented interfaces, so it behaves the same. Workaround: declare the property on the model class itself.

### C# 14 extension properties

Extension properties compile to static methods on the extension container class, not as instance properties of the target type. `typeof(Customer).GetProperties(...)` does not return them, and `ITypeSymbol.GetMembers()` on the target returns only members declared on the type symbol. Neither the runtime walkers nor `ShapeBuilder` see them. Workaround: declare the property on the model class itself — the model already pays the partial-class / Parchment-reference cost, so helper / computed properties belong there.


## Benchmarks

``` ini
BenchmarkDotNet v0.15.8, Windows 11 (10.0.26200.8328)
AMD Ryzen 9 5900X 3.70GHz, 1 CPU, 24 logical and 12 physical cores
.NET SDK 11.0.100-preview.3.26207.106
  [Host] : .NET 10.0.7 (10.0.7, 10.0.726.21808), X64 RyuJIT x86-64-v3
```


### Registration

| Method | Mean | Allocated |
|---|---|---|
| RegisterFromMemoryStream | 0.25 ms | 193 KB |
| RegisterFromBufferedStream | 0.25 ms | 191 KB |
| RegisterFromFilePath | 0.33 ms | 191 KB |


### Rendering

ItemCount varies the number of loop iterations (line items for docx, findings/actions for markdown).

| Method | ItemCount | Mean | Allocated |
|---|---|---|---|
| DocxTemplate | 3 | 0.24 ms | 180 KB |
| MarkdownTemplate | 3 | 0.54 ms | 372 KB |
| DocxTemplate | 50 | 0.41 ms | 295 KB |
| MarkdownTemplate | 50 | 1.04 ms | 771 KB |
| DocxTemplate | 500 | 1.98 ms | 1,444 KB |
| MarkdownTemplate | 500 | 6.86 ms | 4,549 KB |


### Loops

A pure-loop docx template (`{% for line in Lines %}{{ line.Description }}: {{ line.Quantity }} x {{ line.UnitPrice }}{% endfor %}`) rendered against varying iteration counts.

| Method | LoopItems | Mean | Allocated |
|---|---|---|---|
| RenderLoop | 10 | 0.21 ms | 181 KB |
| RenderLoop | 100 | 0.55 ms | 409 KB |
| RenderLoop | 1000 | 3.85 ms | 2,669 KB |

Run benchmarks with:

```
dotnet run --project src/Parchment.Benchmarks --configuration Release
```


## Determinism

Rendering the same template with the same model produces a byte-identical output. Useful for hash-based caching, dedup, and legal sign-off workflows.

[Editable fields](#editable-fields-two-way-binding) preserve the guarantee: content-control and editable-range ids are assigned sequentially in processing order, and document protection is passwordless by design — a password would require a random salt, breaking byte-identical output while adding no real security.


## Icon

[Parchment](https://thenounproject.com/icon/parchment-6992235/) icon designed by Alum Design from [The Noun Project](https://thenounproject.com).
