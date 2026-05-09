# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and test

The repo has two solutions: `src/` (library + source generator + tests + sample model) and `IntegrationTests/` (separate solution that consumes the packed nuget end-to-end).

```bash
# Build everything in src
dotnet build src --configuration Release

# Build the integration tests (requires src to have been built/packed first — Parchment is consumed as a PackageReference, not a ProjectReference)
dotnet build IntegrationTests --configuration Release
```

Tests use **TUnit**, not NUnit/xUnit. Each test project is `OutputType=Exe` with `TestingPlatformDotnetTestSupport=true`, and they are run via `dotnet run`, not `dotnet test`:

```bash
dotnet run --project src/Parchment.Tests --configuration Release
dotnet run --project src/Parchment.SourceGenerator.Tests --configuration Release
dotnet run --project IntegrationTests/IntegrationTests --configuration Release

# Run a single test (TUnit's filter syntax)
dotnet run --project src/Parchment.Tests --configuration Release -- --filter "FullyQualifiedName~Substitution"
```

Snapshot testing uses **Verify.TUnit + Verify.OpenXml + Morph.OpenXml.Skia**. A failing test writes `*.received.*` files alongside the corresponding `*.verified.*`. To accept a new/changed snapshot, rename `received` → `verified` (or use a Verify diff tool). PNG page renders only fire on `net10.0` when `Morph.OpenXml.Skia` is referenced — that's automatic, no opt-in needed.

CI is `src/appveyor.yml` (AppVeyor). The build script first installs every TTF/OTF in `src/Fonts/` into the Windows fonts dir (Aptos is bundled), validating each via `System.Drawing.Text.PrivateFontCollection` so a CRLF-mangled font fails the build instead of producing confusing "font not found" errors at test time.

## Architecture

Parchment combines a .NET data model with one of two template formats and produces a `.docx`. The **two flows** share token scanning, the Fluid singletons, the markdown renderer, and the snapshot/test pipeline, but execute very differently.

### Flow A — Docx template (`RegisterDocxTemplate<T>` / `Render`)

This is the structurally interesting flow. The template is a real `.docx` containing liquid tokens scattered across paragraphs (substitutions like `{{ Customer.Name }}` plus block tags `{% for %}` / `{% if %}`). Handing the whole file to Fluid is not viable, because body paragraphs contain Word run-level formatting that text-based liquid would destroy. Instead:

1. **Token scanner** (`Tokens/TokenScanner.cs`, `Word/ParagraphText.cs`) — walks each paragraph in `MainDocumentPart`, every `HeaderPart`, `FooterPart`, `FootnotesPart`, `EndnotesPart`. For each paragraph it builds a flat `InnerText` plus a `RunMap` that maps character offsets back to the source `<w:t>` elements, regex-scans for `{{ ... }}` and `{% ... %}`, and classifies the paragraph as `Static` / `Substitution` / `Block`.

2. **Anchor bookmarks** (`Word/Anchors.cs`) — every token-bearing paragraph gets an invisible `<w:bookmarkStart w:name="parchment-anchor-{guid}"/>` injected at scan time. These survive byte-stream cloning intact and let render-time look up host paragraphs by name without fragile positional indices. Bookmarks are stripped from the final output before save.

3. **Scope tree** (`Tokens/RangeNode.cs`, `ScopeTreeBuilder.cs`) — once paragraphs are classified, a recursive walk groups block-tag paragraphs with their bodies into `LoopNode` / `IfNode` / `IfBranch` records, with non-block paragraphs as `SubstitutionNode` / `StaticNode` leaves. Block tags MUST sit alone in their own paragraph; mixed inline text + block tag is rejected at registration time.

4. **Reference validator** (`TemplateStore.cs::ReferenceValidator`) — walks the scope tree and, for each token, runs `ModelValidator` against `TModel` via reflection. Loop variables introduce scoped bindings (`item: ElementType`) into a `Dictionary<string, Type>` that nested loops/conditionals inherit. Throws `ParchmentRegistrationException` on the first missing member.

5. **Render time** (`Tokens/ScopeTreeRunner.cs`, `RegisteredDocxTemplate.cs`) — clones the canonical bytes into a fresh `WordprocessingDocument`, builds an anchor → paragraph map, walks the cached scope tree:
   - **Substitutions**: get the parsed `FluidTemplate`, pull its `OutputStatement.Expression`, call `Expression.EvaluateAsync(context)`, check whether the resulting `FluidValue.ToObjectValue()` is a `TokenValue`. If so, queue a structural replacement; otherwise apply in-paragraph text substitution. In-paragraph substitutions are applied **in reverse offset order** so earlier offsets stay valid.
   - **Loops**: call `LoopNode.LoopSource.EvaluateAsync(context).Enumerate(context)` to iterate (this is Fluid's own `Expression`, not a wrapped template). For each item, deep-clone the body elements, rewrite their anchor-bookmark names to fresh per-iteration GUIDs, and recursively process the clones with the loop variable bound in the scope.
   - **Conditionals**: call `IfBranch.Condition.EvaluateAsync(context).ToBooleanValue()`. The chosen branch is processed in place; non-chosen branch paragraphs are removed.
   - Structural replacements are applied last, top-down: each host paragraph is replaced with the produced elements via `parent.InsertAfter` + `host.Remove`.

**Important**: `ScopeTreeRunner` uses Fluid's actual AST nodes (`Fluid.Ast.Expression`, `ForStatement`, `IfStatement`, `OutputStatement`) for runtime evaluation. There is NO hand-rolled loop iteration via reflection, and no string-render of conditions to `"true"`/`"false"`. `IdentifierVisitor` is only used at **registration time** for compile-time-style validation against a .NET `Type`.

### Flow B — Markdown template (`RegisterMarkdownTemplate<T>` / `Render`)

The whole `.md` source is one Fluid template parsed once. At render time:

1. Render the cached `IFluidTemplate` against the model → final markdown text.
2. `Markdig.Markdown.Parse(text, MarkdigPipeline.Pipeline)` → `MarkdownDocument`.
3. Open the optional `styleSource` docx (or a built-in blank), preserve its `<w:sectPr>` (page size, margins, header/footer references), clear the body, walk the markdown AST through `OpenXmlMarkdownRenderer`, append elements, re-attach the sectPr.
4. Header parts, footer parts, styles, theme, font tables, and every other part are passed through verbatim — the renderer only mutates `Body` and (additively) `NumberingDefinitionsPart`.

`MarkdigPipeline` enables: emphasis-extras, grid + pipe tables, autolinks, list-extras (alpha + roman), smarty pants, and **generic attributes** (`{.StyleName}` syntax for attaching Word styles to headings/paragraphs — the renderer honors the first class on each block).

### `OpenXmlMarkdownRenderer` — container stack pattern

The markdown renderer (`Markdown/OpenXmlMarkdownRenderer.cs`) subclasses `Markdig.Renderers.RendererBase` (NOT `TextRendererBase` — it has no `TextWriter`). State is a `Stack<ContainerState>` where each `ContainerState` holds `{ List<OpenXmlElement> Blocks, List<OpenXmlElement> CurrentRuns }`. Inline renderers append runs to `Top.CurrentRuns`; block renderers call `FlushParagraph()` to wrap accumulated runs into a `<w:p>`. Nested constructs (table cells, quote blocks) push/pop new container states. **One renderer instance per render** — do not cache; the stack is mutable.

When the renderer encounters an `HtmlBlock`, it delegates to `OpenXmlHtml.WordHtmlConverter.ToElements(html, mainPart, settings)`, passing `HeadingLevelOffset` derived from the renderer's current heading depth. **OpenXmlHtml ≥ 0.4.0 is required** — the heading-offset parameter was added specifically for Parchment.

`HtmlInline` (e.g. `<em>`, `<strong>`, `<u>`, `<s>`, `<sub>`, `<sup>`, `<br>`) is handled directly by `HtmlInlineRenderer`, NOT through OpenXmlHtml. Markdig delivers inline HTML as separate AST nodes (open tag, literal text, close tag), so the renderer maintains a small stack of active `RunProperties` mutators on `OpenXmlMarkdownRenderer` (`PushInlineHtmlFormat` / `PopInlineHtmlFormat`). `AddRun` applies the active stack to every emitted `Run`, so text between `<em>...</em>` picks up `Italic`, between `<strong>...</strong>` picks up `Bold`, etc. `<br>` becomes a `Break` run. Unknown tags (e.g. `<custom>`) fall through as literal text — preserves the legacy behavior for tags Parchment doesn't model.

Markdown `![alt](url)` images delegate to the same OpenXmlHtml converter — `LinkInlineRenderer.WriteImage` synthesizes a one-shot `<img src="<url>" />` and hands it to `WordHtmlConverter.ToElements`. So data URIs, `file://` URIs, absolute / CWD-relative file paths, and `http(s)://` URLs all flow through `OpenXmlHtml.ImageResolver` regardless of whether the source was a `<img>` tag in an `[Html]` property or a `![]()` token in a markdown template. **Do not** reintroduce a parallel "resolve to data URI in Parchment first" path inside `LinkInlineRenderer` — it duplicates ImageResolver and silently bypasses the active `ImagePolicy`.

### Image policies (`Word/ImagePolicies.cs`)

`TemplateStore` exposes two `init`-only properties — `LocalImages` and `WebImages`, both of type `OpenXmlHtml.ImagePolicy`, both defaulting to `ImagePolicy.AllowAll()`. The store builds an internal `ImagePolicies` record (`(ImagePolicy LocalImages, ImagePolicy WebImages)` plus a `BuildSettings(headingOffset, numberingSession)` helper) and threads it through every `OpenXmlHtml.WordHtmlConverter.ToElements` call site:

- `RegisteredDocxTemplate` → `ScopeTreeRunner` (and every cloned runner inside loops / if branches) — used for `HtmlToken` rendering AND when `ScopeTreeRunner` calls `MarkdownRendering.Render` for `MarkdownToken`.
- `RegisteredMarkdownTemplate` → `MarkdownRendering.Render` → `OpenXmlMarkdownRenderer` (exposed as `ImagePolicies` property), which `HtmlBlockRenderer` and `LinkInlineRenderer` read.

`AllowAll` is the deliberate default because Parchment renders developer-bound content; OpenXmlHtml itself defaults to `Deny()` because *it* has no idea where its HTML came from. Don't chase OpenXmlHtml's default — Parchment's threat model is different.

**Lockstep when adding new OpenXmlHtml call sites**: any new `OpenXmlHtml.WordHtmlConverter.ToElements` invocation MUST consume `ImagePolicies` (via `imagePolicies.BuildSettings(...)` for runtime calls, or `renderer.ImagePolicies.BuildSettings(...)` from inside a markdown renderer). A bare `new HtmlConvertSettings()` silently falls back to OpenXmlHtml's `Deny()` defaults and reintroduces the old alt-text-fallback bug for `<img>` tags.

### Fluid integration (`Liquid/SharedFluid.cs`)

A single static `FluidParser` and `TemplateOptions` shared across all renders (Fluid documents these as thread-safe and recommends per-process singletons). `RegisterModel(Type)` walks the model's reachable type graph and calls `MemberAccessStrategy.Register<T>()` on every POCO type via reflection (`MakeGenericMethod`). This is **load-bearing**: Fluid 2.x's `DefaultMemberAccessStrategy` does NOT auto-discover nested types from a model passed via `TemplateContext(model)` — without the recursive walk, `{{ Customer.Name }}` returns empty, leading to obscure debugging sessions.

Filters (`Liquid/Filters.cs`): `markdown`, `escape_xml`, `bullet_list`, `numbered_list`. Filters returning a `TokenValue` (wrapped in `ObjectValue`) trigger structural replacement at the substitution site — don't replace this with a string-only filter unless also wiring a separate detection path.

### Source generator (`src/Parchment.SourceGenerator/`)

`netstandard2.0;net10.0`, packaged inside `Parchment.nupkg` under `analyzers/dotnet/roslyn5.0/cs` and `analyzers/dotnet/roslyn5.3/cs`. `IIncrementalGenerator` + `ForAttributeWithMetadataName("Parchment.ParchmentTemplateAttribute")`. The same attribute supports **both** docx and markdown templates — the SG branches on the path's extension (`.docx` → docx flow, `.md` / `.markdown` → markdown flow). The pipeline collects `targets`, `docs`, and `markdowns` in three parallel `AdditionalTextsProvider` stages (each cached separately for incrementality), then `Combine`s them and dispatches per target.

**Docx flow** (mirrors runtime Flow A):

1. Resolves `path` against `<AdditionalFiles>` in the consuming project.
2. Reads the docx via `AdditionalText.Path` + `ZipFile.OpenRead` (binary AdditionalFiles pattern — `GetText()` returns null for binary, and direct `File.ReadAllBytes` is banned by RS1035).
3. Unzips with `System.IO.Compression.ZipArchive` and walks `word/document.xml`, `word/header*.xml`, `word/footer*.xml`, `word/footnotes.xml`, `word/endnotes.xml` via `XDocument`.
4. Tokenizes paragraph `InnerText` into `{{ ... }}` / `{% ... %}` sites with a small splitter regex, then hands each token to `Fluid.FluidParser` and walks the resulting AST with an `IdentifierVisitor : Fluid.Ast.AstVisitor` to collect member-access paths — same approach as the runtime library. Fluid.Core and its full transitive closure (Parlot, Microsoft.Extensions.FileProviders.Abstractions, TimeZoneConverter, plus System.Text.Json + Microsoft.Bcl.HashCode on netstandard2.0) are merged into `Parchment.SourceGenerator.dll` at build time via `PackageShader.MsBuild` (marked `Shade="true"` in the SG csproj), so the analyzer ships as a single self-contained DLL under `analyzers/dotnet/roslyn*/cs`. The SG keeps a parallel `TokenScanner.cs` rather than sharing source with the runtime library because the runtime types are internal and depend on `DocumentFormat.OpenXml`, which doesn't work in a netstandard2.0 Roslyn analyzer.
5. Validates references against the model's `ITypeSymbol` (Roslyn semantic model, not reflection).
6. Emits the full diagnostic set (`PARCH001`–`PARCH008`, `PARCH010`, `PARCH011`) and a generated `RegisterWith(TemplateStore store, string? basePath = null)` that calls `store.RegisterDocxTemplate<TModel>(name, path)`. Nested-class targets are wrapped in matching `partial {kind} {name} { ... }` declarations for every link in the enclosing-type chain — `BuildPartialSource` walks `target.EnclosingTypes` (outermost first) emitting wrappers before the target's own partial. PARCH011 fires when any enclosing type is not declared `partial` and the registration helper is *not* emitted in that case (otherwise the user would see CS0260 with no link back to the SG's expectation).

**Markdown flow** (mirrors runtime Flow B):

1. Resolves `path` against `<AdditionalFiles>`.
2. Reads the file via `AdditionalText.GetText(cancel)` and stores the raw text in an equatable `MarkdownData` record. `File.ReadAllText` is **not** used — RS1035 bans it for analyzers, and the canonical Roslyn path lets the SDK reuse cached `SourceText`. The test harness's `PathAdditionalText.GetText` returns real `SourceText` for `.md`/`.markdown` paths (and `null` for `.docx`) so SG behaviour matches real Roslyn.
3. Hands the **whole file** to `FluidParser.TryParse` (one call, not per-token) — markdown templates have no concept of paragraph boundaries, so the token-site scanner used for docx would be the wrong tool. A parse failure becomes a `PARCH006` with the Fluid error message.
4. `MarkdownValidator.Validate` walks the parsed AST directly with loop-scope tracking: descends into `ForStatement` (binding `Identifier` → element FQN via `ShapeResolver.GetElementType`), `IfStatement` (incl. `ElseIfs` and `Else`), `OutputStatement`, and the `for-else` branch. Each `MemberExpression` is collected via a private `ExpressionPathCollector : AstVisitor` and resolved against `target.Shape` honouring the current scope. When a loop source doesn't resolve to an enumerable, the loop variable is bound to the root type for the body walk — wrong but minimises cascade noise on top of the upstream `PARCH001`/`PARCH002` already reported.
5. Emits **only** `PARCH001` (MissingMember) and `PARCH002` (LoopSourceNotEnumerable). `PARCH005` (mixed inline block tag) is deliberately not emitted because Fluid parses the whole markdown file as one template — `Hello {% if x %}World{% endif %}` is legal at runtime. `PARCH007`/`PARCH008` (Excelsior) and `PARCH010` (`[Html]`/`[Markdown]` format) are docx-flow-only — `RegisterMarkdownTemplate` doesn't build `ExcelsiorTableMap`/`FormatMap` and there's no structural-replacement dispatch in the markdown path.
6. Generates a `RegisterWith(TemplateStore store, string? basePath = null, Stream? styleSource = null)` that reads the file at runtime via `File.ReadAllText` and calls `store.RegisterMarkdownTemplate<TModel>(name, markdown, styleSource)`. The optional `styleSource` parameter mirrors the runtime API — production callers usually want a brand docx for styles.

**Lockstep with runtime markdown flow**: the runtime's `RegisterMarkdownTemplate` validates references via a heuristic that skips loop variables by string-matching against the source text. The SG validates the *same* references via a proper AST walk with explicit scope, which is strictly more accurate. So the SG will sometimes catch loop-shadowing bugs the runtime misses — when tightening the runtime's loop-variable detection, also revisit `MarkdownValidator` to make sure the rules still match.

The generator deliberately **does not** embed docx or markdown bytes into the assembly — it generates a runtime helper that reads the file from disk. Users deploy templates alongside their app.

`Parchment.SourceGenerator.csproj` has `<InternalsVisibleTo Include="Parchment.SourceGenerator.Tests"/>` — but signed assemblies don't honor IVT without a public-key match, so the test-facing types (`TokenScanner`, `TokenKind`, `Token`) are `public`. Records use `init`, which requires the `IsExternalInit` polyfill in `Polyfills.cs` for the `netstandard2.0` target.

### Excelsior table dispatch (`[ExcelsiorTable]`)

A hook that lets a `{{ Property }}` substitution resolve to a fully-formatted Word table rendered by [Excelsior](https://github.com/SimonCropp/Excelsior) instead of the default string substitution. Parchment takes a hard `PackageReference` on `Excelsior` (>= 2.3.0); the attribute lives in Parchment (`Parchment.ExcelsiorTableAttribute`), so the dep direction is Parchment → Excelsior.

**Runtime path** (`src/Parchment/Excelsior/`):

1. `TemplateStore.RegisterDocxTemplate<TModel>` calls `ExcelsiorTableMap.Build(typeof(TModel), name)` BEFORE scanning parts. The map recursively walks `TModel`'s property graph, collecting every `[ExcelsiorTable]`-marked collection property under its dotted path from the root (e.g. `"Customer.Lines"`). Each entry stores `(DottedPath, ElementType, Func<object, object?> Getter)` where the getter is a chained closure that walks the corresponding object path with null-short-circuiting at every step.
2. `ExcelsiorTokenValidator.Validate` runs immediately after `ReferenceValidator.ValidateTree`. For each substitution token whose first identifier path matches a map entry, it enforces two rules: (a) the token must sit alone in its host paragraph (offset 0, length == paragraph length, no sibling tokens), and (b) the parsed Fluid `OutputStatement.Expression` must be exactly a `MemberExpression` (no filters, no arithmetic, no literals). Both violations throw `ParchmentRegistrationException`.
3. `RegisteredDocxTemplate` stores the map and passes both it AND the original `object model` to every `ScopeTreeRunner` (including cloned runners inside loops/conditionals).
4. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveExcelsiorTable` FIRST, before normal Fluid evaluation. If the token matches, the helper calls `entry.Getter(rootModel)` to fetch the collection as its original CLR type, then synthesizes a `TokenValue.OpenXml(ctx => [ExcelsiorTableBridge.BuildTable(elementType, data, mainPart)])`. The existing `TokenValue.OpenXml` structural-replacement path handles the rest — which is why the token must sit alone in its paragraph.
5. `ExcelsiorTableBridge` uses `ConcurrentDictionary<Type, BuilderInvoker>` to cache reflection-built delegates per element type. Each invoker calls `new WordTableBuilder<T>(data).Build(mainPart)` via `ConstructorInfo.Invoke` + `MethodInfo.Invoke`, returning a `DocumentFormat.OpenXml.Wordprocessing.Table`. The reflection cost is amortized per element type, not per render.

**SG path** (`src/Parchment.SourceGenerator/`):

1. `ShapeBuilder.HasExcelsiorTableAttribute` matches the attribute by full-qualified name string (`"global::Parchment.ExcelsiorTableAttribute"`) — the SG can't `typeof()` it because it doesn't reference Parchment.dll.
2. `MemberEntry` gained an `IsExcelsiorTable` bool flag. It stays primitive for the incremental pipeline's cacheability.
3. `TokenScanner.ParseSubstitution` sets a new `Token.IsPlainIdentifier` flag via `IsPlainMemberAccess`, which checks whether the parsed template is a single `OutputStatement` wrapping a bare `MemberExpression`.
4. `ShapeResolver.IsExcelsiorTableMember` walks a segment path (honoring loop scope) and returns true if the final member carries the `IsExcelsiorTable` flag.
5. `ParchmentTemplateGenerator.ValidateExcelsiorToken` is called from the Substitution case of `ValidateTokens`. It gates on `IsExcelsiorTableMember`, then emits `PARCH007` when `HasOtherContent` is true and `PARCH008` when `IsPlainIdentifier` is false.

The runtime and SG enforce the same two rules but via different mechanisms: runtime inspects the Fluid AST directly at registration; SG reads a boolean that `TokenScanner` baked in when parsing. Both paths are covered by `ExcelsiorTableTests` (Parchment.Tests) and `ExcelsiorToken_*` tests (Parchment.SourceGenerator.Tests).

### Html / Markdown property dispatch (`[Html]` / `[Markdown]`)

Parallel hook to Excelsior, but for string properties rather than collections: a `string`/`string?` property marked with a user-defined `HtmlAttribute` / `MarkdownAttribute` (detected by type name — Parchment does not ship the attributes) or with `[StringSyntax("html")]` / `[StringSyntax("markdown")]` causes its `{{ Property }}` substitution to be structurally replaced instead of text-substituted. Html runs through `OpenXmlHtml.WordHtmlConverter.ToElements`; markdown runs through the same `MarkdownRendering.Render` used by the markdown-template flow.

**Runtime path** (`src/Parchment/Formats/`):

1. `FormatMap.Build(modelType, name)` walks the model's reachable property graph, mirroring `ExcelsiorTableMap.WalkType` (per-branch `visited` HashSet, same `ShouldDescend` leaf-skipping). For each property it checks `[HtmlAttribute]` / `[MarkdownAttribute]` by `attribute.GetType().Name` and `[StringSyntaxAttribute]` by full type name with `Syntax == "html" | "markdown"`. Enforces string-only, throws on `[Html]+[Markdown]` or `[Html]+[StringSyntax("markdown")]` (and vice versa). Emits `(DottedPath → FormatKind)` entries.
2. `FormatTokenValidator.Validate` runs in `TemplateStore.RegisterDocxTemplate` right after `ExcelsiorTokenValidator`. The only registration-time rule is that the parsed `OutputStatement.Expression` must be a `MemberExpression` (no filters / arithmetic / literals — the format kind is selected by attribute, so a filter chain would be silently ignored). Solo-in-paragraph is **not** required — see "Inline-aware structural replacement" below.
3. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveFormatted` AFTER `TryResolveExcelsiorTable`, before standard Fluid evaluation. It walks the dotted path on `rootModel` via `PropertyInfo.GetValue`, reads the string, and returns `TokenValue.Html(text)` / `TokenValue.Markdown(text)`. Null strings yield empty output.
4. `TokenValue.HtmlToken` is handled alongside `MarkdownToken` / `OpenXmlToken` in `BuildStructuralReplacements`; its source string is passed to `OpenXmlHtml.WordHtmlConverter.ToElements(..., mainPart, new())`. For non-solo tokens, `ProcessSubstitutionAsync` consults `ParagraphSplicer` instead — see below.
5. `ScopeTreeRunner` takes `FormatMap` as a constructor dep and propagates it to cloned runners in loop / if branches — same pattern as `excelsiorTables` and `rootModel`.

**Inline-aware structural replacement** (`src/Parchment/Word/ParagraphSplicer.cs`):

When a `[Html]` / `[Markdown]` token shares its host paragraph with other text or sibling tokens, `ScopeTreeRunner.ApplyNonSoloStructural` chooses one of two paths:

- **Inline splice** — the produced element list is exactly one Paragraph (typical for inline-only HTML like `<b>x</b>`, or single-line markdown). `ParagraphSplicer.SpliceInline` rebuilds the host's children as `[host children before token] + [produced paragraph's children, minus pPr] + [host children after token]`. The host paragraph's own `pPr` is preserved; the produced paragraph's `pPr` is dropped.
- **Split** — the produced element list has multiple block-level elements or a non-paragraph block (a table). `ParagraphSplicer.Split` clones the host twice (preserving `pPr` on each), trims the first to the runs/text before the token offset and the second to the runs/text after `offset+length`, and returns `[before, ...produced, after]` for the caller to insert. The original host is removed via the existing `structuralReplacements` queue.

Two non-solo block-shaped tokens in the same paragraph throw a `ParchmentRenderException` — the splits would overlap and there is no defined composition. The author needs to give one token its own paragraph.

**SG path** (`src/Parchment.SourceGenerator/`):

1. `MemberEntry` has primitive `IsHtml` / `IsMarkdown` bools (sealed record, incremental-pipeline-friendly).
2. `ShapeBuilder.DetectFormat(ISymbol)` matches attribute class name strings (`"HtmlAttribute"`, `"MarkdownAttribute"`) and the `StringSyntaxAttribute` FQN `"global::System.Diagnostics.CodeAnalysis.StringSyntaxAttribute"` — the SG can't `typeof()` them because neither attribute is shipped by Parchment.
3. `ShapeResolver.ResolveMember` returns the final `MemberEntry` for a segment path (honoring loop scope), analogous to `IsExcelsiorTableMember` but returning the whole entry so the caller can consult both flags without walking the shape twice.
4. `ParchmentTemplateGenerator.ValidateFormatToken` gates on `member.IsHtml || member.IsMarkdown` and emits `PARCH010` on `!IsPlainIdentifier`. **PARCH009 (the legacy "must sit alone" diagnostic) is retired** — the runtime now splices inline / splits the host paragraph, so non-solo tokens are valid.

**Runtime + SG lockstep**: both still enforce plain-member-access (`PARCH010` / runtime `RequirePlainIdentifier`). The relevant tests are `FormatAttributeTests` (runtime, including `NonSoloHtml_*` / `NonSoloMarkdown_*` for the inline-splice and split paths) and `FormatToken_*` (SG).

Loop-scoped tokens fall through — `FormatMap` is keyed on dotted paths from the root model only, matching the `ExcelsiorTableMap` limitation. `{% for line in Lines %}{{ line.Body }}{% endfor %}` where `Line.Body` is `[Html]` won't trigger structural replacement; use `{{ line.Body | markdown }}` / the `markdown` filter explicitly inside loops.

### String-list dispatch (auto bullet list for `IEnumerable<string>`)

Mirrors Excelsior's "Enumerable string properties" feature: a substitution token whose dotted path on the root model resolves to an `IEnumerable<string>` (including `string[]`, `List<string>`, `IReadOnlyList<string>`, etc.) auto-renders as a Word native bullet list — same output as `{{ Tags | bullet_list }}` produces explicitly. **No attribute is required**; detection is purely type-driven. There is no SG counterpart and no diagnostic codes — the path is opportunistic, not validated.

**Runtime path** (`src/Parchment/StringLists/`):

1. `StringListMap.Build(modelType, name)` walks the model's reachable property graph. For each property:
   - Skip if `[ExcelsiorTable]` is present (Excelsior keeps ownership; the attribute permits `string` element types and a `[ExcelsiorTable] IEnumerable<string>` would otherwise be shadowed).
   - Otherwise, if the property type is assignable to `IEnumerable<string>` (excluding `string` itself, which is `IEnumerable<char>`), add an entry `(DottedPath, Func<object, object?> Getter)` keyed on the dotted path. Same `ChainGetter` / `ShouldDescend` / per-branch `visited` discipline as `ExcelsiorTableMap`.
2. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveStringList` AFTER `TryResolveExcelsiorTable` and `TryResolveFormatted`, before standard Fluid evaluation. Dispatch order: Excelsior → Format → StringList → Fluid.
3. `TryResolveStringList` **gates opportunistically** instead of throwing on misuse (the design choice that distinguishes this from Excelsior/Format paths):
   - Returns `null` if the token isn't solo in its paragraph (sibling count != 1, or offset/length doesn't cover the whole paragraph text).
   - Returns `null` if the parsed `OutputStatement.Expression` isn't a plain `MemberExpression` (i.e. user attached a filter — they're explicitly opting into Fluid-driven rendering).
   - In both fall-through cases, Fluid takes over. This preserves backward compat: `{{ Tags | bullet_list }}` and `{{ Tags | numbered_list }}` keep working unchanged.
4. When the gates pass, the helper calls `entry.Getter(rootModel)`, materializes the sequence with `.ToList()` (so the deferred render delegate doesn't re-enumerate), and returns `TokenValueHelpers.BulletList(items)` — the same primitive the `bullet_list` filter uses, which produces a `TokenValue.OpenXmlToken` wrapping `IOpenXmlContext.CreateBulletNumbering()` + `ListParagraph`-styled paragraphs.
5. `ScopeTreeRunner` takes `StringListMap` as a constructor dep and propagates it to cloned runners in loop / if branches — same pattern as `excelsiorTables` and `formats`.

**Why opportunistic, not strict** (different from Excelsior/Format): there's an existing `bullet_list` Liquid filter that takes any enumerable and produces the same Word bullet output. Pre-feature, users wrote `{{ Tags | bullet_list }}` against `IEnumerable<string>` properties. A strict validator would have broken that — see `TokenOverrideTests.BulletListFilter` against `Invoice.Tags`. The fall-through design lets the new feature be purely additive: `{{ Tags }}` solo → auto bullet, `{{ Tags | numbered_list }}` → user-driven numbered list, `{{ Tags }}` mixed inline → Fluid stringification (unchanged from before).

Loop-scoped tokens fall through — `StringListMap` is keyed on dotted paths from the root model only, matching the `ExcelsiorTableMap` and `FormatMap` limitations. Inside `{% for c in Customers %}{{ c.Tags }}{% endfor %}` use the explicit `bullet_list` filter.

The relevant test file is `StringListTests` (Parchment.Tests/Docx). The scenario lives at `src/Parchment.Tests/Scenarios/string-list/`.

### Determinism guarantee

Same template + same model → byte-identical output. Avoid `w:rsid` randomness, never set `PackageProperties.Created`, no timestamps anywhere. `DeterminismTests.cs` renders a sample twice and asserts byte equality. Users hash outputs for caching/dedup, so don't break this.

### Scenario directories (`src/Parchment.Tests/Scenarios/`)

Self-contained example folders used by the readme to show before/after of a feature. One subdirectory per scenario (e.g. `Scenarios/excelsior-table/`), each holding every artifact needed to illustrate the feature. Layout:

```
src/Parchment.Tests/Scenarios/
├── ScenarioInputRenderer.cs          # [Explicit] — regenerates all input.png files
└── <scenario-name>/
    ├── input.docx                    # committed binary — the template
    ├── input.png                     # "before" render of input.docx (generated)
    ├── output.verified.docx          # Verify snapshot of the rendered output
    ├── output#page01.verified.png    # "after" render of output.verified.docx (from Verify.OpenXml + Morph)
    ├── output#00.verified.txt        # Verify text extraction
    └── output#01.verified.txt
```

**Why the pattern exists**: so the readme can reference `scenarios/<name>/input.png` and `scenarios/<name>/output#page01.verified.png` as a clean before/after pair without scattering the images across the test tree. Every file in the directory belongs to that one example.

**How a scenario test is wired up** (see `ExcelsiorTableTests.Render`):

1. The `.cs` test file lives under `src/Parchment.Tests/Docx/` (or wherever feature tests live), not inside the scenario directory — the scenario dir is an asset folder, not a code folder.
2. The test reads `input.docx` from disk via a `[CallerFilePath]`-anchored helper:
   ```csharp
   static string SourcePath([CallerFilePath] string path = "") => path;
   static string ScenarioPath(string name) =>
       Path.GetFullPath(Path.Combine(Path.GetDirectoryName(SourcePath()) ?? "", "..", "Scenarios", name));
   ```
3. It calls `File.ReadAllBytesAsync(Path.Combine(ScenarioPath("..."), "input.docx"))`, registers, renders.
4. It directs Verify's output into the scenario dir with a custom filename prefix:
   ```csharp
   var settings = new VerifySettings();
   settings.UseDirectory(ScenarioPath("excelsior-table"));
   settings.UseFileName("output");
   await Verify(stream, "docx", settings);
   ```
   The `UseFileName("output")` keeps the snapshot files clean — `output.verified.docx` / `output#page01.verified.png` rather than the usual `ClassName.MethodName.*` naming.

**How `input.png` is generated** (`ScenarioInputRenderer.cs`):

Marked `[Test, Explicit]` so it does NOT run in the default `dotnet run` path (83-test Parchment.Tests run). It globs `Scenarios/**/input.docx`, runs each through Morph's SkiaSharp renderer, and writes the first page's PNG bytes next to the source docx:

```csharp
var converter = new WordRender.Skia.DocumentConverter();
var options = new WordRender.ConversionOptions();
var pages = converter.ConvertToImageData(stream, options); // IReadOnlyList<byte[]>
File.WriteAllBytesAsync(pngPath, pages[0]);
```

Invoke it on demand when a scenario's template changes:

```bash
dotnet run --project src/Parchment.Tests --configuration Release -- \
    --treenode-filter "/*/*/ScenarioInputRenderer/RenderAllInputDocxesToPng"
```

TUnit's `[Explicit]` attribute excludes the test from default runs; only an explicit filter targeting it will execute it. This is why `input.png` is *committed* alongside `input.docx` — the explicit test regenerates it on demand, but the committed file is what the readme references.

**Adding a new scenario**:

1. `mkdir src/Parchment.Tests/Scenarios/<name>` and drop `input.docx` in it (via a one-shot generator test, or hand-authored in Word).
2. Mark the binary: `*.docx binary` is already in `.gitattributes`.
3. Write a feature test that loads `input.docx`, renders it, and calls `UseDirectory(ScenarioPath("<name>")) + UseFileName("output")` so the snapshot lands in the scenario dir.
4. Run the explicit `ScenarioInputRenderer` test once to produce `input.png`.
5. Reference both PNGs from `readme.md` with `/src/Parchment.Tests/Scenarios/<name>/input.png` and `/src/Parchment.Tests/Scenarios/<name>/output%23page01.verified.png` (note the `#` → `%23` URL escape).

## Non-obvious gotchas

- **Tokens straddling run boundaries**: Word splits text into multiple `<w:r>` elements when formatting changes, when proofing markers fire, or when smart-quote autocorrect runs. `{{ customer.name }}` can land across N runs. The scanner uses `paragraph.InnerText` + a `RunMap` (offset → `<w:t>` element) so substitutions land correctly. The formatting of the **first run** containing the opening `{{` wins for the entire substitution — document this constraint when adding tests.

- **PascalCase tokens**: Liquid in Parchment uses PascalCase (`{{ Customer.Name }}`), not snake_case. Fluid's default member access compares case-insensitively against the actual property name. There is no snake-case → PascalCase translation layer; an earlier attempt to wire `MemberNameStrategies.SnakeCase` was abandoned because that API doesn't exist in Fluid 2.15.

- **`ScopeTreeRunner.ProcessLoopAsync` attaches each iteration's clones to a scratch `Body` before running the nested scope tree**. Without this, nested `{% for %}` / `{% if %}` silently no-op: `open.Parent` and `NextSibling()` return null on a detached clone, so `CaptureBetween(open, close)` captures nothing and `open.Remove()` does nothing, and the inner block-tag paragraph text lands in the output as literal `{% for ... %}`. Reverting this to `parent.InsertAfter(clone, insertAnchor)` for each clone *before* the nested run breaks nested loops in a way that only the `LoopTests.NestedLoop` test catches.

- **`OpenXmlMarkdownRenderer` is not thread-safe** — one instance per render. The `Stack<ContainerState>` and `ObjectRenderers` collection are mutable. The `RegisteredTemplate` (cached canonical bytes + scope tree) IS immutable and safe to share, so concurrent renders work — each render gets its own renderer.

- **`appveyor.yml` font validation step** — every TTF/OTF in `src/Fonts/` is loaded through `System.Drawing.Text.PrivateFontCollection` BEFORE being copied to `%WINDIR%\Fonts`. This catches Git CRLF corruption upfront. When adding a font, mark it as binary in `.gitattributes` (`*.ttf binary`, `*.otf binary` — already present).

- **`ParchmentModel` is a separate project** (not `Model`) to avoid name clashes with common test fixture names in IDE autocomplete.

- **Excelsior dispatch bypasses Fluid, deliberately**: `ExcelsiorTableBridge` walks the CLR model directly via a cached `Func<object, object?>` getter chain, NOT via `Expression.EvaluateAsync` on the parsed token. Routing through Fluid *looks* tempting — it would "enable filters" — but Fluid's `ArrayValue.ToObjectValue()` returns `FluidValue[]`, which erases the `IEnumerable<T>` type that `new WordTableBuilder<T>(data)` needs. This is why `ScopeTreeRunner` takes `rootModel` as a separate constructor parameter instead of pulling the model from `TemplateContext` — the context's collection values have already been wrapped. If someone "simplifies" this by removing the `rootModel` parameter, the Excelsior path breaks on the first filtered or loop-nested token.

- **Per-branch visited set in `ExcelsiorTableMap.WalkType`**: cycle prevention uses a `HashSet<Type>` that's mutated with `visited.Add` on descend and `visited.Remove` on return. The same type can appear at multiple unrelated paths (e.g. `Order.Buyer.Addresses` + `Order.Seller.Addresses` — both walked into Buyer's/Seller's type), but a self-reference (`Node.Next` → `Node`) is pruned. Converting this to a global visited set that never removes entries would silently drop the second sibling branch — tests may still pass if only one branch is exercised, but registration would start missing `[ExcelsiorTable]` properties in reachable-twice models.

- **SG `[ExcelsiorTable]` detection matches by full-qualified-name string** (`"global::Parchment.ExcelsiorTableAttribute"` in `ShapeBuilder.HasExcelsiorTableAttribute`). The SG can't `typeof()` the attribute because it doesn't reference Parchment.dll (the SG runs inside Roslyn, the runtime library doesn't ship into the analyzer). Renaming or moving the attribute silently breaks `PARCH007`/`PARCH008` until the string literal is updated — there's no compile-time safety net.

- **Excelsior runtime and SG validators must stay in lockstep**: `ExcelsiorTokenValidator` (runtime) and `ValidateExcelsiorToken` + `ShapeResolver.IsExcelsiorTableMember` (SG) enforce the same two rules — solo-in-paragraph and plain-member-access. The runtime checks the Fluid AST directly (`output.Expression is MemberExpression`); the SG piggybacks on a `Token.IsPlainIdentifier` bool set by `TokenScanner.IsPlainMemberAccess`. When tightening or loosening one rule, update the other in the same PR. The relevant tests are `ExcelsiorTableTests` (runtime) and `ExcelsiorToken_*` (SG).

- **`MemberEntry.IsExcelsiorTable` must stay primitive** — it's a `bool` on a `sealed record` that flows through the incremental generator pipeline, so its equality is structural. Adding a `bool` was safe; adding a `List<T>`, an `ISymbol` reference, or any mutable field would defeat the pipeline's cacheability and force ShapeBuilder to re-run on every compilation.

- **Package dependency direction is Parchment → Excelsior, hard** (not the other way around). The `[ExcelsiorTable]` attribute deliberately lives in Parchment, not Excelsior. Moving it to Excelsior would invert the dep: every Excel-only Excelsior consumer would pull Parchment, and every Parchment user would have to reference an attribute-only stub package. Don't.

- **SG markdown reads via `AdditionalText.GetText`, not `File.ReadAllText`**: RS1035 bans `File.*` in analyzers. The markdown branch (`ParchmentTemplateGenerator.ReadMarkdown`) calls `text.GetText(cancel)` and treats a null result as a read error. The test harness's `PathAdditionalText.GetText` returns real `SourceText` for `.md`/`.markdown` paths but `null` for `.docx` (which is read via `ZipFile.OpenRead(text.Path)` in the docx branch). If a future test driver returns `null` for everything, the markdown SG path will surface a `PARCH006` "AdditionalText returned no SourceText" instead of the expected validation diagnostics — the failure mode is loud, not silent.

- **`MarkdownValidator` binds the loop variable to the root type when the source is unresolved**: in `WalkFor`, a loop whose `Source` resolves to nothing (PARCH001) or to a non-enumerable (PARCH002) still binds `Identifier` → `target.Shape.RootTypeFullyQualifiedName` for the body walk. This is intentional cascade-suppression: it stops every body access (`{{ line.Description }}` etc.) from also tripping PARCH001 on top of the upstream diagnostic. Skipping the body walk on bad sources would stop validating refs in nested constructs entirely; binding to a sentinel "unknown" type would generate false positives. The current behaviour is wrong but optimal — keep it.
