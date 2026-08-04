# Parchment.Morph

Resolves [Parchment](https://github.com/Papyrine/Parchment)'s table-of-contents and page-reference
fields by laying the document out with [Morph](https://github.com/Papyrine/Morph).

A page number is a product of layout, so ordinarily a generated document leaves those fields for
Word to compute — which costs the reader a prompt on open and shows placeholder text until they
accept it. This measures the document as it is built and writes the numbers in, so the file opens
with the values already in place.

```csharp
var store = new TemplateStore
{
    PageNumbers = new MorphPageNumberResolver()
};
store.RegisterMarkdownTemplate<ReportModel>("report", markdown, styleSource);
```

The template asks for the fields the same way it always does:

```markdown
Contents {.TOCHeading}

[TOC]{levels=1}

# Summary {#summary}

The summary is on page [](#summary).
```

Two things to weigh. It costs a layout pass per render — pages are measured rather than drawn, but
the document is still laid out in full. And Morph is a very close approximation of Word's layout
rather than Word itself, so a document the two paginate differently ends up with page numbers that
are wrong and, because nothing is left marked for update, look authoritative. The fields survive, so
a reader can still refresh them; nobody refreshes numbers that look right.
