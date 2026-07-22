# Todo

Still-open issues found while migrating [LegislationManager](../../LegislationManager) off Aspose.Words
— 16 Word generators, all using `RegisterMarkdownTemplate` with the existing `.dotx` files as style
sources.

**Verified against `3.1.0`.** An earlier version of this list was written against `3.1.0-beta.10` while
reading newer local HEAD source, so it reported several things that were already fixed. This list is
re-tested: everything below still reproduces on the released `3.1.0`.

For the record, upgrading `beta.10` → `3.1.0` was a strict improvement — **31 snapshots went back to
matching the original Aspose output exactly, with no regressions.** The `<br>` fix (`AddBreakRun`)
closed the single biggest fidelity gap; it was the one recurring difference across every generator and
it affected real database content, not only authored markup.

## Fixed in this working tree

 * **Empty paragraphs.** `OpenXmlHtml` dropped any paragraph with no runs, so `<p></p>` was a no-op and
   a blank spacer line could not be written as html. An explicitly empty text block now emits an empty
   paragraph, keeping any paragraph properties on it. Containers are excluded, and a trailing bare one
   is still trimmed.
 * **Word form fields.** `RegisterDocxTemplate` now rewrites legacy `FORMTEXT` fields into
   `{{ Name }}` tokens at registration, so a template authored as a Word form binds unchanged. This
   removes the need for the one-off `overview.dotx` conversion the migration had to script.
 * **`.dotx` form templates.** `EnsureDocumentType` ran after the parts were scanned, and
   `ChangeDocumentType` swaps the main part out — so the rewrite was discarded along with the part it
   was made on, and a form saved as a `.dotx` silently kept its `FORMTEXT` fields. The type change now
   runs before anything reads the parts.

## Open

### 1. A Word tab is still unreachable from an html block

`Emit word tabs from markdown (#50)` covers markdown text. But content inside a raw html block still
goes through the html path, where a literal tab collapses to a space — correct html, but it means
`<w:tab/>` cannot be produced from an html block at all.

This matters because raw html blocks are the only way to express several things (cell merges, shading,
per-cell rich content), so a table that needs both `colspan` **and** tab-aligned text has no route.

Reproduced on `3.1.0`: a literal tab inside `<p>cc\tFirst Parliamentary Counsel</p>` renders as a space.
Worked around with a private-use sentinel rewritten to `<w:tab/>` in a post-render pass.

Suggested: honour `&#9;` (or a `white-space: pre` span) in the html path, so the escape hatch exists.

### 2. `font-weight: normal` cannot override a bold paragraph style

`OpenXmlHtml/src/OpenXmlHtml/HtmlParser.cs` maps it to `format.Bold = false`, but the Word writer
encodes false as the **absence** of `<w:b/>` rather than an explicit `<w:b w:val="0"/>`.

Input: `<h3><b>SMITH</b><span style="font-weight: normal">, John</span></h3>`

Output:

```xml
<w:r><w:rPr><w:b/></w:rPr><w:t>SMITH</w:t></w:r><w:r><w:t>, John</w:t></w:r>
```

The second run has no `rPr` at all, so under `Heading3` (bold) it inherits bold. Aspose's
`builder.Bold = false` emitted the explicit off. No html-side workaround exists.

### 3. Percentage cell widths are silently dropped

| markup | result |
|---|---|
| `<td width="35%">` | no `<w:tcW>` emitted at all |
| `<td style="width:35%">` | no `<w:tcW>` emitted at all |
| `<td style="width:250px">` | `<w:tcW w:w="3750" w:type="dxa"/>` |

The readme documents both the `width` attribute and cell css `width` as supported, and Word has
`w:type="pct"`, so percentages are expressible. Forcing px means hand-computing twips against the page
width — fragile, and wrong the moment margins change.

### 4. A single-cell table collapses without a `<td>` width

`<table style="width:602px">` correctly emits `<w:tblW w:w="9030" w:type="dxa"/>`, but the cell still
shrinks to its content unless the `<td>` also carries a px width. Surprising given the table width was
accepted.

### 5. Whitespace is not folded the way a browser folds it

`<p>Line1\r\n\r\nLine2</p>` keeps the doubled spaces. Aspose's html import collapsed runs of whitespace
like a browser does, so db-sourced html needs a manual `\s+` → `" "` pass plus stripping whitespace
after `<p>`.

### 6. A `<li>` containing a block-level `<p>` loses its bullet

`<li><p>x</p></li>` renders unbulleted while `<li>x</li>` bullets correctly. Description html that
happens to be `<p>`-wrapped silently stops being a list — worth either unwrapping or keeping the
numbering.

### 7. An empty `<div style="page-break-before: always">` emits no break

There is no paragraph for the property to attach to, so nothing happens. The css has to hang off a real
paragraph (e.g. the heading that starts the new page). Reasonable once known, but it fails silently.

Related, and arguably a Verify.OpenXml issue rather than this repo's: the text extractor reports
`--- Page Break ---` for `<w:br w:type="page"/>` but not for `<w:pageBreakBefore/>`, so pagination that
is actually correct looks like a lost break in snapshots.

## Docs

### 8. `src/Parchment/nuget-readme.md` is stale

Shows `RegisterDocxTemplate<Invoice>("invoice", File.ReadAllBytes(...))` and
`var bytes = await store.Render("invoice", model);` — neither overload exists. The root `readme.md` is
correct.

### 9. Readme claims not matched by behaviour

Cell `width` documented as supported, percentages dropped (#3).

## Theme: failures are silent

`3.1.0` improved this markedly — loop-body validation and the static-member warning both landed, and both
would have saved real debugging time here.

What remains follows the same shape: **a plausible-looking document with content quietly missing or
unstyled, and no diagnostic.** Items #1, #2, #3, #6 and #7 above each produce a valid document that is
wrong, and each was caught only because there was a snapshot to diff against. A consumer without
one ships the broken document.

A diagnostic (or debug log) on a dropped css property or an ignored attribute would close the remaining
gap.
