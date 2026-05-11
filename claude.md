# CLAUDE.md

Guidance for Claude Code working in this repo.

## Build and test

Two solutions: `src/` (library + source generator + tests + sample model) and `IntegrationTests/` (consumes the packed nuget end-to-end).

```bash
dotnet build src --configuration Release
dotnet build IntegrationTests --configuration Release   # requires src to be built/packed first — Parchment is a PackageReference, not a ProjectReference
```

Tests use **TUnit** (not NUnit/xUnit). Each test project is `OutputType=Exe` with `TestingPlatformDotnetTestSupport=true` — run via `dotnet run`, not `dotnet test`:

```bash
dotnet run --project src/Parchment.Tests --configuration Release
dotnet run --project src/Parchment.SourceGenerator.Tests --configuration Release
dotnet run --project IntegrationTests/IntegrationTests --configuration Release

# Single test (TUnit filter syntax)
dotnet run --project src/Parchment.Tests --configuration Release -- --filter "FullyQualifiedName~Substitution"
```

Snapshots: **Verify.TUnit + Verify.OpenXml + Morph.OpenXml.Skia**. Failed test writes `*.received.*`; accept by renaming to `*.verified.*`. PNG page renders fire on `net10.0` automatically when `Morph.OpenXml.Skia` is referenced.

CI: `src/appveyor.yml`. Build first installs every TTF/OTF in `src/Fonts/` into the Windows fonts dir, validating each via `System.Drawing.Text.PrivateFontCollection` so a CRLF-mangled font fails the build instead of producing "font not found" at test time.

## Architecture

Parchment combines a .NET data model with one of two template formats and produces a `.docx`. The two flows share token scanning, the Fluid singletons, the markdown renderer, and the snapshot pipeline, but execute very differently.

### Flow A — Docx template (`RegisterDocxTemplate<T>` / `Render`)

Template is a real `.docx` containing liquid tokens scattered across paragraphs. Handing the whole file to Fluid is not viable — body paragraphs contain Word run-level formatting that text-based liquid would destroy. Instead:

1. **Token scanner** (`Tokens/TokenScanner.cs`, `Word/ParagraphText.cs`) — walks paragraphs in `MainDocumentPart`, every `HeaderPart`, `FooterPart`, `FootnotesPart`, `EndnotesPart`. For each paragraph builds flat `InnerText` + a `RunMap` (offset → `<w:t>` element), regex-scans for `{{ ... }}` and `{% ... %}`, classifies as `Static` / `Substitution` / `Block`.
2. **Anchor bookmarks** (`Word/Anchors.cs`) — every token-bearing paragraph gets an invisible `<w:bookmarkStart w:name="parchment-anchor-{guid}"/>` at scan time. These survive byte-stream cloning so render-time can look up host paragraphs by name. Stripped before save.
3. **Scope tree** (`Tokens/RangeNode.cs`, `ScopeTreeBuilder.cs`) — recursive walk groups block-tag paragraphs with their bodies into `LoopNode` / `IfNode` / `IfBranch`, with non-block paragraphs as `SubstitutionNode` / `StaticNode` leaves. Block tags MUST sit alone in their paragraph; mixed inline text + block tag is rejected at registration.
4. **Reference validator** (`TemplateStore.cs::ReferenceValidator`) — walks scope tree, runs `ModelValidator` against `TModel` via reflection. Loop variables introduce scoped bindings (`item: ElementType`) into a `Dictionary<string, Type>` that nested loops/conditionals inherit. Throws `ParchmentRegistrationException` on first missing member.
5. **Render** (`Tokens/ScopeTreeRunner.cs`, `RegisteredDocxTemplate.cs`) — clones canonical bytes, builds anchor → paragraph map, walks the cached scope tree:
   - **Substitutions**: get parsed `FluidTemplate`, pull `OutputStatement.Expression`, call `Expression.EvaluateAsync(context)`. If result's `ToObjectValue()` is a `TokenValue`, queue structural replacement; else apply in-paragraph text substitution. In-paragraph substitutions applied **in reverse offset order** so earlier offsets stay valid.
   - **Loops**: `LoopNode.LoopSource.EvaluateAsync(context).Enumerate(context)` (Fluid's own `Expression`, not a wrapped template). For each item, deep-clone body elements, rewrite anchor-bookmark names to fresh per-iteration GUIDs, recursively process the clones with loop variable bound in scope.
   - **Conditionals**: `IfBranch.Condition.EvaluateAsync(context).ToBooleanValue()`. Chosen branch processed in place; non-chosen branch paragraphs removed.
   - Structural replacements applied last, top-down, via `parent.InsertAfter` + `host.Remove`.

**Important**: `ScopeTreeRunner` uses Fluid's actual AST nodes (`Fluid.Ast.Expression`, `ForStatement`, `IfStatement`, `OutputStatement`) at runtime. NO hand-rolled loop iteration via reflection, no string-render of conditions to `"true"`/`"false"`. `IdentifierVisitor` is only for **registration-time** validation against a .NET `Type`.

### Flow B — Markdown template (`RegisterMarkdownTemplate<T>` / `Render`)

Whole `.md` source is one Fluid template parsed once. At render:

1. Render cached `IFluidTemplate` against model → final markdown text.
2. `Markdig.Markdown.Parse(text, MarkdigPipeline.Pipeline)` → `MarkdownDocument`.
3. Open optional `styleSource` docx (or built-in blank), preserve its `<w:sectPr>` (page size, margins, header/footer refs), clear body, walk markdown AST through `OpenXmlMarkdownRenderer`, append, re-attach sectPr.
4. Header/footer parts, styles, theme, font tables passed through verbatim — renderer only mutates `Body` and (additively) `NumberingDefinitionsPart`.

`MarkdigPipeline` enables: emphasis-extras, grid + pipe tables, autolinks, list-extras (alpha + roman), smarty pants, **generic attributes** (`{.StyleName}` for attaching Word styles to headings/paragraphs — renderer honors first class on each block).

### `OpenXmlMarkdownRenderer` — container stack pattern

`Markdown/OpenXmlMarkdownRenderer.cs` subclasses `Markdig.Renderers.RendererBase` (NOT `TextRendererBase` — no `TextWriter`). State is `Stack<ContainerState>` where each holds `{ List<OpenXmlElement> Blocks, List<OpenXmlElement> CurrentRuns }`. Inline renderers append runs to `Top.CurrentRuns`; block renderers call `FlushParagraph()` to wrap accumulated runs into `<w:p>`. Nested constructs (table cells, quote blocks) push/pop new container states. **One renderer per render** — do not cache; the stack is mutable.

`HtmlBlock` delegates to `OpenXmlHtml.WordHtmlConverter.ToElements(html, mainPart, settings)`, passing `HeadingLevelOffset` from current heading depth. **OpenXmlHtml ≥ 0.4.0 required** — heading-offset parameter was added for Parchment.

`HtmlInline` (`<em>`, `<strong>`, `<u>`, `<s>`, `<sub>`, `<sup>`, `<br>`) handled directly by `HtmlInlineRenderer`, NOT through OpenXmlHtml. Markdig delivers inline HTML as separate AST nodes (open tag, literal text, close tag), so renderer maintains a stack of active `RunProperties` mutators (`PushInlineHtmlFormat`/`PopInlineHtmlFormat`). `AddRun` applies the active stack to every emitted `Run`. `<br>` becomes a `Break` run. Unknown tags fall through as literal text.

Markdown `![alt](url)` images delegate to OpenXmlHtml — `LinkInlineRenderer.WriteImage` synthesizes a one-shot `<img src="<url>" />` and hands to `WordHtmlConverter.ToElements`. So data URIs, `file://`, absolute/CWD-relative paths, and `http(s)://` all flow through `OpenXmlHtml.ImageResolver` regardless of source. **Do not** reintroduce a parallel "resolve to data URI in Parchment first" path inside `LinkInlineRenderer` — duplicates ImageResolver and bypasses `ImagePolicy`.

### Image policies (`Word/ImagePolicies.cs`)

`TemplateStore` exposes `init`-only `LocalImages` and `WebImages` (`OpenXmlHtml.ImagePolicy`), both defaulting to `ImagePolicy.AllowAll()`. The store builds an internal `ImagePolicies` record (`(LocalImages, WebImages)` + `BuildSettings(headingOffset, numberingSession)`) and threads it through every `OpenXmlHtml.WordHtmlConverter.ToElements` call site:

- `RegisteredDocxTemplate` → `ScopeTreeRunner` (and cloned runners in loops/if branches) — for `HtmlToken` AND when `ScopeTreeRunner` calls `MarkdownRendering.Render` for `MarkdownToken`.
- `RegisteredMarkdownTemplate` → `MarkdownRendering.Render` → `OpenXmlMarkdownRenderer` (exposed as `ImagePolicies` property), read by `HtmlBlockRenderer` and `LinkInlineRenderer`.

`AllowAll` is the deliberate default because Parchment renders developer-bound content; OpenXmlHtml defaults to `Deny()` because *it* doesn't know its HTML's source. Don't chase OpenXmlHtml's default — Parchment's threat model is different.

**Lockstep**: any new `WordHtmlConverter.ToElements` call MUST consume `ImagePolicies` (via `imagePolicies.BuildSettings(...)` for runtime, or `renderer.ImagePolicies.BuildSettings(...)` from inside a markdown renderer). A bare `new HtmlConvertSettings()` falls back to OpenXmlHtml's `Deny()` and reintroduces the alt-text-fallback bug for `<img>`.

### Fluid integration (`Liquid/SharedFluid.cs`)

Single static `FluidParser` and `TemplateOptions` shared across all renders (Fluid documents these as thread-safe; recommends per-process singletons). `RegisterModel(Type)` walks model's reachable type graph and calls `MemberAccessStrategy.Register<T>()` on every POCO via reflection. **Load-bearing**: Fluid 2.x's `DefaultMemberAccessStrategy` does NOT auto-discover nested types from a model passed via `TemplateContext(model)` — without the recursive walk, `{{ Customer.Name }}` returns empty.

Filters (`Liquid/Filters.cs`): `markdown`, `escape_xml`, `bullet_list`, `numbered_list`. Filters returning a `TokenValue` (wrapped in `ObjectValue`) trigger structural replacement at the substitution site.

### Source generator (`src/Parchment.SourceGenerator/`)

`netstandard2.0;net10.0`, packaged inside `Parchment.nupkg` under `analyzers/dotnet/roslyn5.0/cs` and `analyzers/dotnet/roslyn5.3/cs`. `IIncrementalGenerator` + `ForAttributeWithMetadataName("Parchment.ParchmentModelAttribute")`. Same attribute supports **both** docx and markdown — SG branches on extension (`.docx` → docx flow, `.md` → markdown flow). The attribute target IS the binding model (no separate marker class — see "Design decisions" below). Pipeline collects `targets`, `docs`, `markdowns` in three parallel `AdditionalTextsProvider` stages (each cached separately for incrementality), `Combine`s them, dispatches per target.

**Docx flow** (mirrors runtime Flow A):

1. Resolves `path` against `<AdditionalFiles>`.
2. Reads docx via `AdditionalText.Path` + `ZipFile.OpenRead` (binary AdditionalFiles pattern — `GetText()` returns null for binary, and `File.ReadAllBytes` is banned by RS1035).
3. Walks `word/document.xml`, `word/header*.xml`, `word/footer*.xml`, `word/footnotes.xml`, `word/endnotes.xml` via `XDocument`.
4. Tokenizes paragraph `InnerText` into `{{ ... }}` / `{% ... %}` sites, hands each to `Fluid.FluidParser` and walks the AST with `IdentifierVisitor : Fluid.Ast.AstVisitor` to collect member-access paths. Fluid.Core and full transitive closure (Parlot, Microsoft.Extensions.FileProviders.Abstractions, TimeZoneConverter, plus System.Text.Json + Microsoft.Bcl.HashCode on netstandard2.0) merged into `Parchment.SourceGenerator.dll` at build via `PackageShader.MsBuild` (marked `Shade="true"`), so analyzer ships as a single self-contained DLL. SG keeps a parallel `TokenScanner.cs` rather than sharing source — runtime types are internal and depend on `DocumentFormat.OpenXml`, which doesn't work in a netstandard2.0 analyzer.
5. Validates against model's `ITypeSymbol` (Roslyn semantic model, not reflection).
6. Emits diagnostics `PARCH001`–`PARCH008`, `PARCH010`, `PARCH011`. Generates `RegisterWith(TemplateStore store, string? basePath = null)` calling `store.RegisterDocxTemplate<TModel>(name, path)`. Nested-class targets are wrapped in matching `partial {kind} {name}` declarations for every link in the enclosing-type chain — `BuildPartialSource` walks `target.EnclosingTypes` (outermost first). PARCH011 fires when any enclosing type isn't `partial`, and the helper is *not* emitted in that case (otherwise CS0260 with no link back to the SG).
7. **Emits pre-compiled accessors** via `AccessorEmission.Emit(target.Shape, rootFqn)`: a `static readonly KeyValuePair<string, IMemberAccessor>[]` block per non-empty type (one `DelegateAccessor` per member) plus `static readonly ExcelsiorTableMapEntry[]` / `FormatMapEntry[]` / `StringListMapEntry[]` arrays for the per-template maps. `RegisterWith` calls `Parchment.Generated.GeneratedRegistration.RegisterFluidAccessors` / `RegisterExcelsiorTable` / `RegisterFormat` / `RegisterStringList` for each before delegating to `store.RegisterDocxTemplate<TModel>`. The runtime maps (`SharedFluid.registeredTypes`, `ExcelsiorTableMap.precompiledCache`, etc.) see every type already populated and short-circuit reflection. Getters for nested dotted paths are emitted as `o => ((Root)o).First?.Second?.Third` — first segment uses `.` (root is non-null per Render contract), subsequent segments use `?.` to match the runtime ChainGetter's null-short-circuit semantics. Non-nullable value-type *intermediates* would break the chain — accepted limitation since binding-model intermediates are virtually always reference types.

**Markdown flow** (mirrors runtime Flow B):

1. Resolves `path` against `<AdditionalFiles>`.
2. Reads via `AdditionalText.GetText(cancel)` into an equatable `MarkdownData` record. `File.ReadAllText` is **not** used — RS1035 bans it; canonical Roslyn path lets the SDK reuse cached `SourceText`. Test harness's `PathAdditionalText.GetText` returns real `SourceText` for `.md` (and `null` for `.docx`).
3. Hands the **whole file** to `FluidParser.TryParse` (one call, not per-token) — markdown templates have no paragraph boundaries. Parse failure → `PARCH006` with Fluid error.
4. `MarkdownValidator.Validate` walks the AST with loop-scope tracking: descends into `ForStatement` (binding `Identifier` → element FQN via `ShapeResolver.GetElementType`), `IfStatement` (incl. `ElseIfs`/`Else`), `OutputStatement`, `for-else`. Each `MemberExpression` is collected via private `ExpressionPathCollector : AstVisitor` and resolved against `target.Shape` honouring scope. When loop source doesn't resolve to enumerable, loop var is bound to root type for body walk — minimises cascade noise on top of upstream `PARCH001`/`PARCH002`.
5. Emits **only** `PARCH001` (MissingMember) and `PARCH002` (LoopSourceNotEnumerable). `PARCH005` (mixed inline block tag) deliberately not emitted — Fluid parses whole markdown as one template; `Hello {% if x %}World{% endif %}` is legal at runtime. `PARCH007`/`PARCH008` (Excelsior) and `PARCH010` (`[Html]`/`[Markdown]`) are docx-only — `RegisterMarkdownTemplate` doesn't build `ExcelsiorTableMap`/`FormatMap`.
6. Generates `RegisterWith(TemplateStore store, string? basePath = null, Stream? styleSource = null)` reading the file at runtime via `File.ReadAllText` and calling `store.RegisterMarkdownTemplate<TModel>(name, markdown, styleSource)`. Same `AccessorEmission` pass as docx — the markdown registration helper also pre-compiles every Fluid accessor + map entry, even though markdown templates don't currently use `[ExcelsiorTable]` / `[Html]` / `[Markdown]` dispatch (those are docx-only). Symmetric emission keeps the SG output predictable regardless of which template shape is being registered.

**Lockstep with runtime markdown flow**: runtime's `RegisterMarkdownTemplate` validates references via heuristic that skips loop variables by string-matching against source text. SG validates the *same* references via proper AST walk with explicit scope — strictly more accurate. So SG sometimes catches loop-shadowing bugs the runtime misses — when tightening runtime's loop-variable detection, also revisit `MarkdownValidator`.

The generator deliberately **does not** embed docx/markdown bytes — generates a runtime helper that reads from disk. Users deploy templates alongside their app.

`Parchment.SourceGenerator.csproj` has `<InternalsVisibleTo Include="Parchment.SourceGenerator.Tests"/>` — but signed assemblies don't honor IVT without public-key match, so test-facing types (`TokenScanner`, `TokenKind`, `Token`) are `public`. Records use `init`, requiring the `IsExternalInit` polyfill in `Polyfills.cs` for netstandard2.0.

### Excelsior table dispatch (`[ExcelsiorTable]`)

A hook letting `{{ Property }}` resolve to a fully-formatted Word table rendered by [Excelsior](https://github.com/SimonCropp/Excelsior) instead of default string substitution. Parchment takes a hard `PackageReference` on `Excelsior` (>= 2.3.0); the attribute lives in Parchment (`Parchment.ExcelsiorTableAttribute`), so dep direction is Parchment → Excelsior.

**Runtime path** (`src/Parchment/Excelsior/`):

1. `TemplateStore.RegisterDocxTemplate<TModel>` calls `ExcelsiorTableMap.Build(typeof(TModel), name)` BEFORE scanning parts. Recursively walks `TModel`'s property graph, collecting every `[ExcelsiorTable]`-marked collection under its dotted path (e.g. `"Customer.Lines"`). Each entry stores `(DottedPath, ElementType, Func<object, object?> Getter)` — getter is a chained closure with null-short-circuiting at every step.
2. `ExcelsiorTokenValidator.Validate` runs immediately after `ReferenceValidator.ValidateTree`. For each substitution token whose first identifier matches a map entry: (a) token must sit alone in its host paragraph, (b) parsed `OutputStatement.Expression` must be exactly a `MemberExpression` (no filters/arithmetic/literals). Both throw `ParchmentRegistrationException`.
3. `RegisteredDocxTemplate` stores the map and passes both it AND original `object model` to every `ScopeTreeRunner` (including cloned runners in loops/conditionals).
4. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveExcelsiorTable` FIRST. If matched, calls `entry.Getter(rootModel)`, synthesizes `TokenValue.OpenXml(ctx => [ExcelsiorTableBridge.BuildTable(elementType, data, mainPart)])`. Existing `TokenValue.OpenXml` structural-replacement path handles the rest — which is why the token must sit alone.
5. `ExcelsiorTableBridge` uses `ConcurrentDictionary<Type, BuilderInvoker>` to cache reflection-built delegates per element type. Each invoker calls `new WordTableBuilder<T>(data).Build(mainPart)` via `ConstructorInfo.Invoke` + `MethodInfo.Invoke`. Reflection cost amortized per element type.

**SG path** (`src/Parchment.SourceGenerator/`):

1. `ShapeBuilder.HasExcelsiorTableAttribute` matches by FQN string (`"global::Parchment.ExcelsiorTableAttribute"`) — SG can't `typeof()` it (doesn't reference Parchment.dll).
2. `MemberEntry` has `IsExcelsiorTable` bool flag — primitive for incremental pipeline cacheability.
3. `TokenScanner.ParseSubstitution` sets `Token.IsPlainIdentifier` via `IsPlainMemberAccess` (single `OutputStatement` wrapping a bare `MemberExpression`).
4. `ShapeResolver.IsExcelsiorTableMember` walks segment path (honoring loop scope), returns true if final member carries the flag.
5. `ParchmentTemplateGenerator.ValidateExcelsiorToken` (in Substitution case of `ValidateTokens`) gates on `IsExcelsiorTableMember`, emits `PARCH007` when `HasOtherContent` and `PARCH008` when `!IsPlainIdentifier`.

Runtime and SG enforce same two rules via different mechanisms: runtime inspects Fluid AST directly; SG reads a boolean baked in by `TokenScanner`. Tests: `ExcelsiorTableTests` (runtime) and `ExcelsiorToken_*` (SG).

### Html / Markdown property dispatch (`[Html]` / `[Markdown]`)

Parallel to Excelsior but for string properties: a `string`/`string?` property marked with user-defined `HtmlAttribute`/`MarkdownAttribute` (detected by type name — Parchment doesn't ship the attributes) or `[StringSyntax("html")]`/`[StringSyntax("markdown")]` causes its `{{ Property }}` to be structurally replaced. Html runs through `OpenXmlHtml.WordHtmlConverter.ToElements`; markdown runs through the same `MarkdownRendering.Render` used by markdown-template flow.

**Runtime path** (`src/Parchment/Formats/`):

1. `FormatMap.Build(modelType, name)` walks property graph mirroring `ExcelsiorTableMap.WalkType` (per-branch `visited` HashSet, same `ShouldDescend` leaf-skipping). Checks `[HtmlAttribute]`/`[MarkdownAttribute]` by `attribute.GetType().Name` and `[StringSyntaxAttribute]` by full type name with `Syntax == "html" | "markdown"`. Enforces string-only; throws on `[Html]+[Markdown]` or `[Html]+[StringSyntax("markdown")]` (and vice versa).
2. `FormatTokenValidator.Validate` runs in `RegisterDocxTemplate` right after `ExcelsiorTokenValidator`. Only registration-time rule: parsed `OutputStatement.Expression` must be a `MemberExpression`. Solo-in-paragraph is **not** required — see splice below.
3. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveFormatted` AFTER `TryResolveExcelsiorTable`, before standard Fluid eval. Walks dotted path on `rootModel` via `PropertyInfo.GetValue`, returns `TokenValue.Html(text)`/`TokenValue.Markdown(text)`. Null strings → empty.
4. `TokenValue.HtmlToken` handled alongside `MarkdownToken`/`OpenXmlToken` in `BuildStructuralReplacements`; source string passed to `OpenXmlHtml.WordHtmlConverter.ToElements(..., mainPart, new())`. For non-solo tokens, `ProcessSubstitutionAsync` consults `ParagraphSplicer`.
5. `ScopeTreeRunner` takes `FormatMap` as ctor dep, propagates to cloned runners.

**Inline-aware structural replacement** (`src/Parchment/Word/ParagraphSplicer.cs`):

When `[Html]`/`[Markdown]` token shares its host paragraph with other text/tokens, `ScopeTreeRunner.ApplyNonSoloStructural` chooses:

- **Inline splice** — produced list is exactly one Paragraph (typical for inline-only HTML like `<b>x</b>`, or single-line markdown). `ParagraphSplicer.SpliceInline` rebuilds host's children as `[before token] + [produced paragraph's children, minus pPr] + [after token]`. Host's `pPr` preserved; produced `pPr` dropped.
- **Split** — produced list has multiple block-level elements or a non-paragraph (a table). `ParagraphSplicer.Split` clones the host twice (preserving `pPr` on each), trims first to runs/text before the offset, second to runs/text after `offset+length`, returns `[before, ...produced, after]`.

Two non-solo block-shaped tokens in the same paragraph throw `ParchmentRenderException` — splits would overlap, no defined composition. Author needs to give one its own paragraph.

**SG path** (`src/Parchment.SourceGenerator/`):

1. `MemberEntry` has primitive `IsHtml`/`IsMarkdown` bools.
2. `ShapeBuilder.DetectFormat(ISymbol)` matches attribute class name strings (`"HtmlAttribute"`, `"MarkdownAttribute"`) and `StringSyntaxAttribute` FQN `"global::System.Diagnostics.CodeAnalysis.StringSyntaxAttribute"`.
3. `ShapeResolver.ResolveMember` returns the final `MemberEntry` for a segment path (honoring loop scope), so caller can consult both flags without walking shape twice.
4. `ParchmentTemplateGenerator.ValidateFormatToken` gates on `member.IsHtml || member.IsMarkdown`, emits `PARCH010` on `!IsPlainIdentifier`. **PARCH009 (legacy "must sit alone") is retired** — runtime now splices/splits.

**Lockstep**: both still enforce plain-member-access (`PARCH010` / runtime `RequirePlainIdentifier`). Tests: `FormatAttributeTests` (runtime, including `NonSoloHtml_*`/`NonSoloMarkdown_*`) and `FormatToken_*` (SG).

Loop-scoped tokens fall through — `FormatMap` is keyed on dotted paths from root model only (same as `ExcelsiorTableMap`). `{% for line in Lines %}{{ line.Body }}{% endfor %}` where `Line.Body` is `[Html]` won't trigger structural replacement — use `{{ line.Body | markdown }}` explicitly inside loops.

### String-list dispatch (auto bullet for `IEnumerable<string>`)

Mirrors Excelsior's "Enumerable string properties": a substitution token whose dotted path resolves to `IEnumerable<string>` (incl. `string[]`, `List<string>`, `IReadOnlyList<string>`, etc.) auto-renders as a Word bullet list — same output as `{{ Tags | bullet_list }}`. **No attribute required**; detection is type-driven. No SG counterpart, no diagnostics — opportunistic, not validated.

**Runtime path** (`src/Parchment/StringLists/`):

1. `StringListMap.Build(modelType, name)` walks property graph. For each property: skip if `[ExcelsiorTable]` (Excelsior keeps ownership; attribute permits `string` element types and would otherwise be shadowed). Else if assignable to `IEnumerable<string>` (excluding `string` itself, which is `IEnumerable<char>`), add `(DottedPath, Func<object, object?> Getter)`. Same `ChainGetter`/`ShouldDescend`/per-branch `visited` discipline as `ExcelsiorTableMap`.
2. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveStringList` AFTER `TryResolveExcelsiorTable` and `TryResolveFormatted`. Dispatch order: Excelsior → Format → StringList → Fluid.
3. `TryResolveStringList` **gates opportunistically** instead of throwing on misuse:
   - Returns `null` if token isn't solo in its paragraph.
   - Returns `null` if parsed `OutputStatement.Expression` isn't a plain `MemberExpression` (user attached a filter — explicitly opting into Fluid).
   - In both cases, Fluid takes over — preserves backward compat.
4. When gates pass: calls `entry.Getter(rootModel)`, materializes with `.ToList()` (so deferred render delegate doesn't re-enumerate), returns `TokenValueHelpers.BulletList(items)` — same primitive the `bullet_list` filter uses (`TokenValue.OpenXmlToken` wrapping `IOpenXmlContext.CreateBulletNumbering()` + `ListParagraph` paragraphs).
5. `ScopeTreeRunner` takes `StringListMap` as ctor dep, propagates to cloned runners.

**Why opportunistic, not strict**: there's an existing `bullet_list` filter producing the same output. Pre-feature, users wrote `{{ Tags | bullet_list }}` against `IEnumerable<string>` properties — strict validator would have broken that (see `TokenOverrideTests.BulletListFilter` against `Invoice.Tags`). Fall-through design keeps the new feature purely additive: `{{ Tags }}` solo → auto bullet, `{{ Tags | numbered_list }}` → user-driven, `{{ Tags }}` mixed inline → Fluid stringification (unchanged).

Loop-scoped tokens fall through — same root-only key limitation as `ExcelsiorTableMap`/`FormatMap`. Inside `{% for c in Customers %}{{ c.Tags }}{% endfor %}` use the explicit `bullet_list` filter.

Tests: `StringListTests` (Parchment.Tests/Docx). Scenario: `src/Parchment.Tests/Scenarios/string-list/`.

### Determinism guarantee

Same template + same model → byte-identical output. Avoid `w:rsid` randomness, never set `PackageProperties.Created`, no timestamps. `DeterminismTests.cs` renders a sample twice and asserts byte equality. Users hash outputs for caching/dedup — don't break this.

### Scenario directories (`src/Parchment.Tests/Scenarios/`)

Self-contained example folders the readme references for before/after of a feature. One subdirectory per scenario:

```
src/Parchment.Tests/Scenarios/
├── ScenarioInputRenderer.cs          # [Explicit] — regenerates all input.png files
└── <scenario-name>/
    ├── input.docx                    # committed binary — the template
    ├── input.png                     # "before" render (generated)
    ├── output.verified.docx          # Verify snapshot
    ├── output#page01.verified.png    # "after" render (Verify.OpenXml + Morph)
    ├── output#00.verified.txt        # Verify text extraction
    └── output#01.verified.txt
```

**Test wiring** (see `ExcelsiorTableTests.Render`):

1. `.cs` file lives under `src/Parchment.Tests/Docx/` (or wherever feature tests live), not inside the scenario dir — scenario dir is asset-only.
2. Reads `input.docx` via a `[CallerFilePath]`-anchored helper:
   ```csharp
   static string SourcePath([CallerFilePath] string path = "") => path;
   static string ScenarioPath(string name) =>
       Path.GetFullPath(Path.Combine(Path.GetDirectoryName(SourcePath()) ?? "", "..", "Scenarios", name));
   ```
3. Directs Verify into the scenario dir with a custom prefix:
   ```csharp
   var settings = new VerifySettings();
   settings.UseDirectory(ScenarioPath("excelsior-table"));
   settings.UseFileName("output");
   await Verify(stream, "docx", settings);
   ```
   `UseFileName("output")` → `output.verified.docx` rather than `ClassName.MethodName.*`.

**Regenerating `input.png`**: `ScenarioInputRenderer.cs` is `[Test, Explicit]` — excluded from default runs. It globs `Scenarios/**/input.docx`, renders each via Morph/SkiaSharp, writes the first page next to the source. Invoke on demand:

```bash
dotnet run --project src/Parchment.Tests --configuration Release -- \
    --treenode-filter "/*/*/ScenarioInputRenderer/RenderAllInputDocxesToPng"
```

`input.png` is *committed* alongside `input.docx` — the explicit test regenerates it; the committed file is what the readme references.

**Adding a scenario**: mkdir under `Scenarios/`, drop `input.docx` (binary already in `.gitattributes`), write a feature test that calls `UseDirectory(ScenarioPath("<name>")) + UseFileName("output")`, run `ScenarioInputRenderer` once for `input.png`, reference both PNGs from `readme.md` (note `#` → `%23` URL escape).

## Design decisions

### Source generator emits pre-compiled accessors; runtime reflection is a fallback

The SG path is the **primary** registration path; the runtime reflection walks in `SharedFluid.RegisterTypeGraph`, `ExcelsiorTableMap.WalkType`, `FormatMap.WalkType`, and `StringListMap.WalkType` are kept only as a fallback for callers that can't use the SG. The pattern mirrors `System.Text.Json`'s source generator (`JsonSerializerContext` mode), EF Core's compiled models, and the regex source generator — modern .NET libraries that care about trimming / NativeAOT / cold-start expose a SG that emits the metadata reflection would otherwise compute at runtime, and the runtime checks an SG-populated cache before falling back.

**The pipeline.** `AccessorEmission.Emit(target.Shape, rootFqn)` walks `ModelShape` and produces four parallel datasets:

1. **Fluid accessors**: one `KeyValuePair<string, IMemberAccessor>[]` per non-empty type in the shape (one `DelegateAccessor` per public property / field), registered via `GeneratedRegistration.RegisterFluidAccessors(typeof(T), arr)`. Lands in `SharedFluid.registeredTypes` + `Options.MemberAccessStrategy.Register(type, accessors)`, so the reflection walk in `RegisterTypeGraph` short-circuits at every node.
2. **Excelsior-table entries**: dotted-path walk from the root collecting `[ExcelsiorTable]` members. Each entry is `(DottedPath, ElementType, Func<object, object?> Getter)`. Registered via `GeneratedRegistration.RegisterExcelsiorTable(typeof(TRoot), arr)` → populates `ExcelsiorTableMap.precompiledCache`.
3. **Format entries**: same walk pattern for `[Html]` / `[Markdown]` / `[StringSyntax]` string properties. Each entry is `(DottedPath, FormatMapKind, Getter)`. Both the registration walk AND the render-time `ScopeTreeRunner.TryResolveFormatted` use the emitted getter — render-time `walker.GetType().GetProperty(...)` reflection per segment is gone on the SG path.
4. **String-list entries**: dotted-path walk for `IEnumerable<string>` members that aren't already owned by `[ExcelsiorTable]`. Each entry is `(DottedPath, Getter)`. Populates `StringListMap.precompiledCache`.

All four runtime maps have `static readonly ConcurrentDictionary<Type, X> precompiledCache` checked at the top of `Build(modelType, ...)`. A cache hit returns the SG-built map and skips the reflection walk entirely. A miss falls through to the existing reflection path (the runtime API contract is unchanged for non-SG callers).

**Public API surface.** SG-emitted code calls into the `Parchment.Generated` namespace:

- `Parchment.Generated.GeneratedRegistration` — static class with one `Register*` method per dataset.
- `Parchment.Generated.ExcelsiorTableMapEntry`, `FormatMapEntry`, `StringListMapEntry` — public data carrier records.
- `Parchment.Generated.FormatMapKind` — public enum (replaces the internal `FormatKind` that previously lived in `Formats/FormatMap.cs`; the internal duplicate was deleted).

These types must stay public — the SG runs against user assemblies referencing Parchment via NuGet, so `[InternalsVisibleTo]` won't help (signed assembly + arbitrary user PKs). The rest of the runtime (`SharedFluid`, the maps themselves, `RegisterTypeGraph`, etc.) stays internal.

**Path-getter emission.** Cast-and-null-chain lambdas: `o => ((global::NS.Root)o).First?.Second?.Third`. The cast on the root is unconditional (`Render` rejects null models). Subsequent segments use `?.` to match the runtime `ChainGetter`'s "null at any level → null result" semantics. Limitation: non-nullable value-type intermediates would fail to compile (`Customer?.Origin` is fine if `Origin` is `Point?`, fails if `Origin` is `Point`). Binding-model intermediates are virtually always reference types in practice, so accepted as a known limitation rather than handled with a manual cascade.

**Why `RegisterPrecompiled` on each map is `internal` (not `public`) but `GeneratedRegistration` is `public`.** The runtime maps are implementation details; the data carrier records + entry point in `Parchment.Generated` are the public surface. This lets the SG cross the assembly boundary without exposing `ExcelsiorTableMap` / `FormatMap` / `StringListMap` / `SharedFluid` themselves (which expose `Fluid.IMemberAccessor`, `IFluidParser`, and other Fluid-coupled types that should not appear in the public API).

**Lockstep concerns.** The SG-side `AccessorEmission.WalkForMaps` and the runtime-side `WalkType` implementations of each map must produce equivalent results — different mechanisms (SG: Roslyn symbols + flag flags baked into `MemberEntry` at extract time; runtime: reflection at registration). Equivalent results means: same dotted-path key set, same element types, same kinds. When tightening one walker, revisit the other. Tests in `Parchment.Tests` exercise the runtime path; SG snapshots in `Parchment.SourceGenerator.Tests` lock in the emitted shape; `IntegrationTests` compile + run the SG output against the packed nupkg.

### `[ParchmentModel]` lives on the binding model, not on an intermediary "template" class

The source-generator attribute is `Parchment.ParchmentModelAttribute` and is applied **directly to the model class being bound** (the type Parchment renders against). There is **no separate marker / "template" class**, and the attribute does not take a `typeof(TModel)` argument — the attribute target *is* the model.

```csharp
[ParchmentModel("Templates/report.md")]
public partial class Report
{
    public string Title { get; set; }

    [Markdown]
    public string Body { get; set; }

    // Helper that adapts a complex graph into a binding-friendly primitive.
    public string FormattedTotal => Total.ToString("C", Culture);
}
```

The SG emits `TemplatePath`, `TemplateName`, and `RegisterWith(TemplateStore store, ...)` into the model partial — the same shape previously emitted onto the marker class.

**Rationale.** Models almost always need code on them anyway:

- `[Html]` / `[Markdown]` annotations on string properties (structural replacement dispatch).
- `[ExcelsiorTable]` on collection properties.
- Helper / computed properties that pre-shape values into binding-friendly form (currency formatting, joined name strings, derived flags).
- Conversions of complex CLR types into the primitives liquid/Fluid can render directly.

Because the model is already a place where the author writes Parchment-aware code, the `partial` requirement and the dependency on Parchment.dll are **already paid**. Adding a separate marker class would force a second declaration site for zero gain — it would not eliminate either tax, and it would add a "where does this live?" decision per template.

**Consequences accepted by this decision:**

- The model **must be `partial`** (the SG generates `RegisterWith` onto it). Models that come from EF/JSON/codegen and resist `partial` are not supported via the SG path — those consumers fall back to the runtime `TemplateStore.RegisterDocxTemplate<T>(name, path)` / `RegisterMarkdownTemplate<T>(name, markdown)` API.
- The model **references `Parchment`** (for the attribute). For most projects this matches reality — the model is already coupled to rendering through `[Html]` / `[Markdown]` / `[ExcelsiorTable]`. Projects wanting a pure POCO model use the runtime API.
- **One template per model via the SG attribute** is the canonical case. Multi-template-per-model scenarios are served by the runtime API, not by stacking attributes — the SG emits exactly one `RegisterWith` per model with no name disambiguation needed.
- The attribute name is `ParchmentModelAttribute`, **not** `ParchmentTemplateAttribute`, to reflect that the decorated type is the model being bound, not a stand-in for the template.

**Alternatives considered and rejected:**

- *Attribute on a separate marker `partial class FooTemplate`* — original design. Rejected: extra declaration with no body, doesn't relieve the `partial` or dependency cost (those move to the model the moment any `[Html]` / `[Markdown]` / helper-property is needed), and forces the user to invent a naming convention for the marker.
- *Assembly-level `[assembly: ParchmentModel(...)]`* — rejected for now: keeps the model POCO, but pushes binding declarations away from the model they describe and forces a separate "registry" namespace per assembly. Discoverability suffers, and the POCO benefit evaporates as soon as the model needs `[Html]` / `[Markdown]`.
- *Supporting multiple placement modes simultaneously* — rejected: triples the SG's attribute-target validation, diagnostic-location, and incremental-pipeline surface for a marginal ergonomics win. One canonical path keeps the SG, the diagnostics, and the docs coherent.

**Implications for future work:**

- Diagnostics referencing "the decorated class" target the model itself. `PARCH011` (enclosing-type must be `partial`) still applies when the model is a nested type.
- The runtime `TemplateStore.RegisterDocxTemplate<T>` / `RegisterMarkdownTemplate<T>` API stays — it is the supported escape hatch for POCO models, multi-template-per-model, or dynamically resolved templates. Do not deprecate it in favor of the SG-only path.
- When touching the SG (`ParchmentTemplateGenerator.cs`), the attribute predicate, target shape, generated partial wrapping (`BuildPartialSource`), and any new diagnostics should all treat the attribute target as the model type. There is no second symbol to thread through.

## Non-obvious gotchas

- **Tokens straddling run boundaries**: Word splits text into multiple `<w:r>` when formatting changes, proofing markers fire, or smart-quote autocorrect runs. `{{ customer.name }}` can land across N runs. Scanner uses `paragraph.InnerText` + `RunMap` (offset → `<w:t>`) so substitutions land correctly. Formatting of the **first run** containing the opening `{{` wins for the entire substitution.

- **PascalCase tokens**: Liquid in Parchment uses PascalCase (`{{ Customer.Name }}`). Fluid's default member access compares case-insensitively. No snake-case → PascalCase translation; an earlier attempt to wire `MemberNameStrategies.SnakeCase` was abandoned because that API doesn't exist in Fluid 2.15.

- **Method-call syntax (`{{ Customer.FullName() }}`) fails at registration, not silently at render**: Fluid's parser rejects function-call syntax by default (`AllowFunctions` is off and stays off — Parchment's binding model is data, not behaviour). `FluidParser.TryParse` returns false with `"Functions are not allowed. To enable the feature use the 'AllowFunctions' option. at (col:line)"`. The docx flow (`TokenScanner.ParseSubstitution`) and markdown flow (`TemplateStore.RegisterMarkdownTemplate`) both surface this as `ParchmentRegistrationException` with the Fluid message embedded; SG markdown surfaces it as `PARCH006` and SG docx as a parse diagnostic from its parallel `TokenScanner`. The failure is loud at registration time — `{{ X.Y() }}` never reaches `ModelValidator` or render. Workaround: expose a computed property (`public string FullName => $"{First} {Last}";`) and write `{{ Customer.FullName }}`.

- **Indexer syntax with a string literal (`{{ Customer['FullName'] }}`) is validated the same as dotted access**: Fluid treats `Customer['FullName']` and `Customer.FullName` as the same member resolution at render time. The three path collectors (`Liquid/IdentifierVisitor.cs`, `Parchment.SourceGenerator/IdentifierVisitor.cs`, `MarkdownValidator.ExpressionPathCollector`) all extract the literal string from an `IndexerSegment { Expression: LiteralExpression { Value.Type: FluidValues.String } }` and treat it as if it were an `IdentifierSegment` — so `Customer['NoSuchMember']` fails registration / fires PARCH001 the same as `Customer.NoSuchMember`. Non-literal indexers (numeric, variable, expression — e.g. `Customer[0]`, `Customer[varName]`) still terminate the path walk: those are dictionary/array access at runtime and cannot be validated statically, and Fluid handles them natively. The shared extractor lives at `Liquid/IdentifierVisitor.TryGetStaticName` (runtime) and `Parchment.SourceGenerator/SegmentNames.TryGetStaticName` (SG, used by both `IdentifierVisitor` and `ExpressionPathCollector`); when tightening one, mirror the other. Side effect: Excelsior / Format / StringList dispatch now also accepts indexer notation because they key on `site.References[0].Dotted` — `{{ Customer['Lines'] }}` dispatches identically to `{{ Customer.Lines }}`. Tests: `ReferenceValidatorTests.Indexer*` (runtime), `ParchmentTemplateGeneratorTests.*IndexerStringLiteral*` (SG).

- **`ScopeTreeRunner.ProcessLoopAsync` attaches each iteration's clones to a scratch `Body` before running the nested scope tree**. Without this, nested `{% for %}`/`{% if %}` silently no-op: `open.Parent` and `NextSibling()` return null on a detached clone, so `CaptureBetween(open, close)` captures nothing, `open.Remove()` does nothing, and the inner block-tag paragraph text lands as literal `{% for ... %}`. Reverting to `parent.InsertAfter(clone, insertAnchor)` for each clone *before* the nested run breaks nested loops in a way only `LoopTests.NestedLoop` catches.

- **`OpenXmlMarkdownRenderer` is not thread-safe** — one instance per render. The `Stack<ContainerState>` and `ObjectRenderers` are mutable. The `RegisteredTemplate` (cached canonical bytes + scope tree) IS immutable — concurrent renders work, each gets its own renderer.

- **`appveyor.yml` font validation** — every TTF/OTF in `src/Fonts/` is loaded through `System.Drawing.Text.PrivateFontCollection` BEFORE being copied to `%WINDIR%\Fonts`, catching Git CRLF corruption upfront. When adding a font, mark it binary in `.gitattributes` (`*.ttf binary`, `*.otf binary` — already present).

- **`ParchmentModel` is a separate project** (not `Model`) to avoid name clashes with common test fixture names in IDE autocomplete.

- **Excelsior dispatch bypasses Fluid, deliberately**: `ExcelsiorTableBridge` walks the CLR model directly via cached `Func<object, object?>` getter chains, NOT via `Expression.EvaluateAsync`. Routing through Fluid *looks* tempting (would "enable filters") but Fluid's `ArrayValue.ToObjectValue()` returns `FluidValue[]`, erasing the `IEnumerable<T>` type that `new WordTableBuilder<T>(data)` needs. This is why `ScopeTreeRunner` takes `rootModel` as a separate ctor param instead of pulling the model from `TemplateContext` — context's collections have already been wrapped. "Simplifying" by removing the `rootModel` param breaks Excelsior on the first filtered or loop-nested token.

- **Per-branch visited set in `ExcelsiorTableMap.WalkType`**: cycle prevention uses a `HashSet<Type>` mutated with `visited.Add` on descend and `visited.Remove` on return. Same type can appear at multiple unrelated paths (e.g. `Order.Buyer.Addresses` + `Order.Seller.Addresses`), but a self-reference (`Node.Next` → `Node`) is pruned. Converting to a global visited set that never removes silently drops the second sibling branch — tests may pass if only one branch is exercised, but registration starts missing `[ExcelsiorTable]` properties in reachable-twice models.

- **SG `[ExcelsiorTable]` detection matches by FQN string** (`"global::Parchment.ExcelsiorTableAttribute"` in `ShapeBuilder.HasExcelsiorTableAttribute`). SG can't `typeof()` it (doesn't reference Parchment.dll). Renaming/moving the attribute silently breaks `PARCH007`/`PARCH008` until the literal is updated — no compile-time safety net.

- **Excelsior runtime and SG validators must stay in lockstep**: `ExcelsiorTokenValidator` (runtime) and `ValidateExcelsiorToken` + `ShapeResolver.IsExcelsiorTableMember` (SG) enforce same two rules — solo-in-paragraph and plain-member-access. Runtime checks Fluid AST directly; SG piggybacks on `Token.IsPlainIdentifier` set by `TokenScanner.IsPlainMemberAccess`. Tighten or loosen both in the same PR. Tests: `ExcelsiorTableTests` (runtime), `ExcelsiorToken_*` (SG).

- **`MemberEntry.IsExcelsiorTable` must stay primitive** — it's a `bool` on a `sealed record` flowing through the incremental generator pipeline; equality is structural. Adding a `bool` was safe; adding `List<T>`, an `ISymbol` reference, or any mutable field defeats cacheability and forces ShapeBuilder to re-run on every compilation.

- **Package dep direction is Parchment → Excelsior, hard**. The `[ExcelsiorTable]` attribute deliberately lives in Parchment, not Excelsior. Moving it to Excelsior would invert the dep: every Excel-only Excelsior consumer would pull Parchment, and every Parchment user would have to reference an attribute-only stub package. Don't.

- **SG markdown reads via `AdditionalText.GetText`, not `File.ReadAllText`**: RS1035 bans `File.*` in analyzers. `ParchmentTemplateGenerator.ReadMarkdown` calls `text.GetText(cancel)` and treats null as a read error. Test harness's `PathAdditionalText.GetText` returns real `SourceText` for `.md`/`.markdown` but `null` for `.docx` (read via `ZipFile.OpenRead(text.Path)` in the docx branch). If a future test driver returns `null` for everything, markdown SG surfaces `PARCH006` "AdditionalText returned no SourceText" — failure mode is loud, not silent.

- **Public fields are bindable at every depth, and the runtime registration walks them explicitly**: `SharedFluid.RegisterTypeGraph` registers properties via Fluid's `MemberAccessStrategyExtensions.Register<T>(strategy)` (properties only), THEN enumerates `GetFields(BindingFlags.Public | BindingFlags.Instance)` and registers each as a `DelegateAccessor` (`Fluid.Accessors.DelegateAccessor`) keyed on field name via `Options.MemberAccessStrategy.Register(Type, IEnumerable<KeyValuePair<string, IMemberAccessor>>)`. Both properties' types AND fields' types are recursed for descent. Without the field-walk pass, top-level field access would still work (because `TemplateContext(model, ..., allowModelMembers: true)` reflects on the root instance directly) but nested-via-field access (`{{ Outer.FieldOnInner }}`) would silently render empty — the inner type wouldn't be registered and Fluid would return nothing. SG-side `ShapeBuilder.TryGetMemberType` already accepts both `IPropertySymbol` and `IFieldSymbol`, and `ModelValidator.ResolveMember` already falls back to `GetField`, so the only place fields were under-supported was the Fluid registration step. If a future Fluid upgrade changes `Register<T>()` to include fields, the explicit pass becomes redundant but harmless — it's the safer side of the cache.

- **Interfaces are not bindable as `TModel`**: `TemplateStore.RegisterDocxTemplate<IFoo>` / `RegisterMarkdownTemplate<IFoo>` is blocked at registration by `GuardBindingModel<TModel>` (`TemplateStore.cs`), throwing `ParchmentRegistrationException` with a message that names the interface and points at concrete types. Root cause: `typeof(IFoo).GetProperties(BindingFlags.Public | BindingFlags.Instance)` returns only members declared directly on `IFoo`, missing inherited base-interface members — every reflection walker (`SharedFluid`, `ExcelsiorTableMap`, `FormatMap`, `StringListMap`, `ModelValidator`) would silently miss them without an explicit `GetInterfaces()` walk. The SG path is blocked one level up by `[AttributeUsage(AttributeTargets.Class)]` on `ParchmentModelAttribute`, producing CS0592 if applied to an interface. **Abstract classes are supported** — `GetProperties` on an abstract class walks the base chain normally, and `Render`'s `IsInstanceOfType` check accepts any concrete subclass instance, so abstract-as-polymorphic-binding-surface is a legitimate pattern. C# constraints can't express "concrete reference type" (`where T : class` permits interfaces because they're reference types; `where T : class, new()` blocks interfaces but also blocks records with primary constructors), so the runtime guard is the chosen mechanism. Tests pinning this live in `InterfaceModelTests`.

- **Default implementations on interfaces, C# 14 extension properties, and explicit interface implementations are not bindable** — all three for the same root reason: they don't surface as public instance properties of the target type. An interface-default property (`interface IDoc { string Header => $"=== {Title} ==="; }`) is only callable through the interface type; `typeof(Foo).GetProperties(...)` on a class implementing `IDoc` returns nothing for `Header` unless the class re-declares it. An extension property (`extension(Customer) { public string FullName => ...; }`) compiles to a static method on the extension container, not an instance property of `Customer`. An explicit interface implementation (`string IFoo.Title { get; }`) compiles as a private member named `IFoo.Title` reachable only via the interface — `BindingFlags.Public` skips it, and the `.`-in-name isn't a valid liquid identifier anyway. Every reflection walker (`SharedFluid.RegisterTypeGraph`, `ExcelsiorTableMap`, `FormatMap`, `StringListMap`, `ModelValidator`) misses all three; SG-side `ITypeSymbol.GetMembers()` returns members declared on the type symbol only, and `ShapeBuilder.BuildEntry` walks `current.BaseType` but not `current.Interfaces` and never visits extension containers. Result: tokens fail validation with PARCH001 (SG) or `ParchmentRegistrationException` (runtime). Workaround: declare a regular public instance property on the model class (for explicit impls, a one-line delegating wrapper: `public string Title => ((IFoo)this).Title;`). Inheritance from a base **class** is supported and tested (`ReferenceValidatorTests.InheritedMember_RegistersAndRenders`, `ShadowedMember_DerivedWins`, `ParchmentTemplateGeneratorTests.InheritedMember_Validates`). Supporting interface-default properties would require walking `GetInterfaces()` in every reflection walker + `current.Interfaces` in `ShapeBuilder`, with dedup rules for diamond cases. Extension properties would additionally require scanning the compilation for extension containers and a Fluid resolver fallback for runtime member-access misses. Explicit interface implementations are a deliberate non-goal beyond cost: the feature exists specifically to *hide* a member from the public surface, so binding to it contradicts intent — and there's no clean liquid identifier without inventing a renaming convention. None are on the roadmap — the model is already where helper/computed properties live by design.

- **`MarkdownValidator` binds the loop variable to the root type when source is unresolved**: in `WalkFor`, a loop whose `Source` resolves to nothing (PARCH001) or to a non-enumerable (PARCH002) still binds `Identifier` → `target.Shape.RootTypeFullyQualifiedName` for the body walk. Intentional cascade-suppression: stops every body access (`{{ line.Description }}` etc.) from also tripping PARCH001. Skipping body walk on bad sources would stop validating refs in nested constructs entirely; binding to a sentinel "unknown" would generate false positives. Current behaviour is wrong but optimal — keep it.
