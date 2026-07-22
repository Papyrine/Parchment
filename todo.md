# Todo

Issues found while migrating [LegislationManager](../../LegislationManager) off Aspose.Words — 20 Word
generators, all but one using `RegisterMarkdownTemplate` with the existing `.dotx` files as style
sources.

**Everything on this list is now fixed.** The list is kept as a record of what the migration surfaced
and where the fixes landed.

For the record, upgrading `3.1.0-beta.10` → `3.1.0` was a strict improvement — **31 snapshots went back
to matching the original Aspose output exactly, with no regressions.** The `<br>` fix (`AddBreakRun`)
closed the single biggest fidelity gap; it was the one recurring difference across every generator and
it affected real database content, not only authored markup.

## Fixed in Parchment

 * **Word form fields.** `RegisterDocxTemplate` now rewrites legacy `FORMTEXT` fields into
   `{{ Name }}` tokens at registration, so a template authored as a Word form binds unchanged. This
   removes the need for the one-off `overview.dotx` conversion the migration had to script.
 * **`.dotx` form templates.** `EnsureDocumentType` ran after the parts were scanned, and
   `ChangeDocumentType` swaps the main part out — so the rewrite was discarded along with the part it
   was made on, and a form saved as a `.dotx` silently kept its `FORMTEXT` fields. The type change now
   runs before anything reads the parts.

## Fixed in OpenXmlHtml

 * **Empty paragraphs.** Any paragraph with no runs was dropped, so `<p></p>` was a no-op and a blank
   spacer line could not be written as html. An explicitly empty text block now emits an empty
   paragraph, keeping any paragraph properties on it. Containers are excluded, and a trailing bare one
   is still trimmed.
 * **Page break placement.** `page-break-before` and `page-break-after` both emitted a standalone empty
   paragraph carrying the break. That left a blank line at the top of every new page, and — because
   renderers collapse an empty paragraph — often lost the break entirely: the affected snapshots
   rendered to *fewer* pages than they should have. The break now lands on the paragraph it breaks
   before, which is the element's own for `before` and the following one for `after`. An empty element
   still gets a paragraph of its own, since that is the whole point of one, and a break onto a table
   still takes an empty paragraph ahead of it because a table has no `w:pageBreakBefore`.
 * **Word tabs from html.** A literal tab collapsed to a space wherever it appeared, so `<w:tab/>` could
   not be produced from html at all. Preserved whitespace now emits real `<w:tab/>` elements, making
   `white-space: pre` the escape hatch. Under the default folding rules a tab is still ordinary
   whitespace, matching a browser.
 * **Bullets on block-wrapped list items.** `<li><p>x</p></li>` rendered the marker on its own line with
   the content stranded underneath, so `<p>`-wrapped description html silently stopped being a list.
   The first block child of an item now continues that item's line. This only ever affected the text
   prefix fallback (`ToParagraphs`, and `ToElements` without a `MainDocumentPart`) — real Word numbering
   already handled it.
 * **`font-weight: normal` over a bold paragraph style.** Now emitted as an explicit `<w:b w:val="false"/>`
   rather than by omitting `<w:b/>`, so a `<span style="font-weight: normal">` inside an `<h3>` actually
   renders unbolded instead of inheriting the style's bold.
 * **Percentage cell widths.** `<td width="35%">` and `<td style="width:35%">` emitted no `w:tcW` at all.
   Both now map to `w:type="pct"`.
 * **Single-cell table widths.** A cell in a table with an absolute width shrank to its content unless
   the `<td>` repeated the width. The table width now reaches the cell.
 * **Whitespace folding.** Runs of whitespace now fold to a single space the way a browser folds them,
   rather than reaching Word verbatim.

## Docs

Both readme drift items are resolved: `src/Parchment/nuget-readme.md` no longer shows overloads that do
not exist, and the OpenXmlHtml readme's cell-width claim matches behaviour now that percentages work.
The OpenXmlHtml readme also gained entries for `white-space`, page break placement, and how a block
child of a list item is treated.

## Theme: failures are silent

This was the one recurring shape across everything above: **a plausible-looking document with content
quietly missing or unstyled, and no diagnostic.** Each item produced a valid document that was wrong,
and each was caught only because there was a snapshot to diff against. A consumer without one ships the
broken document.

The page break item is the sharpest example. It had been wrong long enough that the expected output was
baked into the repo's own snapshots, and the tell was only visible once the page renders were compared:
documents that should have spanned three pages had been rendering as one.

Loop-body validation and the static-member warning both landed during the migration and both would have
saved real debugging time. A diagnostic (or debug log) on a dropped css property or an ignored attribute
would close the rest of the gap.
