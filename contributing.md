# Contributing to Parchment

Architecture, design decisions, and non-obvious gotchas for people working on Parchment itself. User-facing feature docs, PARCH diagnostic codes, the `[ParchmentModel]` placement rationale, model binding limitations, and the determinism guarantee live in `readme.md` — link there rather than restating.

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

Snapshots: **Verify.TUnit + Verify.OpenXml + Morph.Skia**. Failed test writes `*.received.*`; accept by renaming to `*.verified.*`. PNG page renders fire on `net10.0` automatically when `Morph.Skia` is referenced.

CI: `src/appveyor.yml`. Build first installs every TTF/OTF in `src/Fonts/` into the Windows fonts dir, validating each via `System.Drawing.Text.PrivateFontCollection` so a CRLF-mangled font fails the build instead of producing "font not found" at test time.

## Architecture

Parchment combines a .NET data model with one of two template formats and produces a `.docx`. User-facing overview of the two modes is in `readme.md` → "Two modes". The two flows share token scanning, the Fluid singletons, the markdown renderer, and the snapshot pipeline, but execute very differently.

### Flow A — Docx template (`RegisterDocxTemplate<T>` / `Render`)

Template is a real `.docx` containing liquid tokens scattered across paragraphs. Handing the whole file to Fluid is not viable — body paragraphs contain Word run-level formatting that text-based liquid would destroy. Instead:

1. **Token scanner** (`Tokens/TokenScanner.cs`, `Word/ParagraphText.cs`) — walks paragraphs in `MainDocumentPart`, every `HeaderPart`, `FooterPart`, `FootnotesPart`, `EndnotesPart`. For each paragraph builds flat `InnerText` + a `RunMap` (offset → `<w:t>` element), regex-scans for `{{ ... }}` and `{% ... %}`, classifies as `Static` / `Substitution` / `Block`.
2. **Anchor bookmarks** (`Word/Anchors.cs`) — every token-bearing paragraph gets an invisible `<w:bookmarkStart w:name="parchment-anchor-{guid}"/>` at scan time. These survive byte-stream cloning so render-time can look up host paragraphs by name. Stripped before save.
3. **Scope tree** (`Tokens/RangeNode.cs`, `ScopeTreeBuilder.cs`) — recursive walk groups block-tag paragraphs with their bodies into `LoopNode` / `IfNode` / `IfBranch`, with non-block paragraphs as `SubstitutionNode` / `StaticNode` leaves. Block tags MUST sit alone in their paragraph; mixed inline text + block tag is rejected at registration (PARCH005 in SG docx flow).
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

User-facing overview, attribute placement, and PARCH001–PARCH012 are in `readme.md` → "Source generator (recommended)". The notes below cover implementation details needed for navigating the SG code.

`netstandard2.0;net10.0`, packaged inside `Parchment.nupkg` under `analyzers/dotnet/roslyn5.0/cs` and `analyzers/dotnet/roslyn5.3/cs`. `IIncrementalGenerator` + `ForAttributeWithMetadataName("Parchment.ParchmentModelAttribute")`. Same attribute supports **both** docx and markdown — SG branches on extension (`.docx` → docx flow, `.md` → markdown flow). Pipeline collects `targets`, `docs`, `markdowns` in three parallel `AdditionalTextsProvider` stages (each cached separately for incrementality), `Combine`s them, dispatches per target.

**Docx flow** (mirrors runtime Flow A):

1. Resolves `path` against `<AdditionalFiles>`.
2. Reads docx via `AdditionalText.Path` + `ZipFile.OpenRead` (binary AdditionalFiles pattern — `GetText()` returns null for binary, and `File.ReadAllBytes` is banned by RS1035).
3. Walks `word/document.xml`, `word/header*.xml`, `word/footer*.xml`, `word/footnotes.xml`, `word/endnotes.xml` via `XDocument`. Also probes `word/settings.xml` for `<w:removePersonalInformation/>` — absence emits `PARCH012` (warning). The runtime clones the template's `settings.xml` into the output, so the warning surfaces a PII-leak path at compile time.
4. Tokenizes paragraph `InnerText` into `{{ ... }}` / `{% ... %}` sites, hands each to `Fluid.FluidParser` and walks the AST with `IdentifierVisitor : Fluid.Ast.AstVisitor` to collect member-access paths. Fluid.Core + transitive closure (Parlot, Microsoft.Extensions.FileProviders.Abstractions, TimeZoneConverter, plus System.Text.Json + Microsoft.Bcl.HashCode on netstandard2.0) merged into `Parchment.SourceGenerator.dll` at build via `PackageShader.MsBuild` (`Shade="true"`), so analyzer ships as a single self-contained DLL. SG keeps a parallel `TokenScanner.cs` rather than sharing source — runtime types are internal and depend on `DocumentFormat.OpenXml`, which doesn't work in a netstandard2.0 analyzer.
5. Validates against model's `ITypeSymbol` (Roslyn semantic model, not reflection).
6. Emits diagnostics `PARCH001`–`PARCH008`, `PARCH010`–`PARCH012`. Generates `RegisterWith(TemplateStore store, string? basePath = null)` calling `store.RegisterDocxTemplate<TModel>(name, path)`. Nested-class targets are wrapped in matching `partial {kind} {name}` declarations for every link in the enclosing-type chain — `BuildPartialSource` walks `target.EnclosingTypes` (outermost first). PARCH011 fires when any enclosing type isn't `partial`, and the helper is *not* emitted in that case (otherwise CS0260 with no link back to the SG).
7. **Emits pre-compiled accessors** via `AccessorEmission.Emit(target.Shape, rootFqn)` — see "Design decisions → Source generator emits pre-compiled accessors" below for the four-dataset pipeline.

**Markdown flow** (mirrors runtime Flow B):

1. Resolves `path` against `<AdditionalFiles>`.
2. Reads via `AdditionalText.GetText(cancel)` into an equatable `MarkdownData` record. `File.ReadAllText` is **not** used — RS1035 bans it; canonical Roslyn path lets the SDK reuse cached `SourceText`. Test harness's `PathAdditionalText.GetText` returns real `SourceText` for `.md` (and `null` for `.docx`).
3. Hands the **whole file** to `FluidParser.TryParse` (one call, not per-token) — markdown templates have no paragraph boundaries. Parse failure → `PARCH006` with Fluid error.
4. `MarkdownValidator.Validate` walks the AST with loop-scope tracking: descends into `ForStatement` (binding `Identifier` → element FQN via `ShapeResolver.GetElementType`), `IfStatement` (incl. `ElseIfs`/`Else`), `OutputStatement`, `for-else`. Each `MemberExpression` is collected via private `ExpressionPathCollector : AstVisitor` and resolved against `target.Shape` honouring scope. When loop source doesn't resolve to enumerable, loop var is bound to root type for body walk — cascade-noise suppression.
5. Emits **only** `PARCH001` and `PARCH002`. `PARCH005` deliberately not emitted — Fluid parses whole markdown as one template; `Hello {% if x %}World{% endif %}` is legal at runtime. `PARCH007`/`PARCH008` (Excelsior) and `PARCH010` (`[Html]`/`[Markdown]`) are docx-only — `RegisterMarkdownTemplate` doesn't build `ExcelsiorTableMap`/`FormatMap`.
6. Generates `RegisterWith(TemplateStore store, string? basePath = null, Stream? styleSource = null)` reading the file at runtime via `File.ReadAllText` and calling `store.RegisterMarkdownTemplate<TModel>(name, markdown, styleSource)`. Same `AccessorEmission` pass as docx — symmetric emission keeps SG output predictable regardless of template shape.

**Lockstep with runtime markdown flow**: runtime's `RegisterMarkdownTemplate` validates references via heuristic that skips loop variables by string-matching against source text. SG validates the *same* references via proper AST walk with explicit scope — strictly more accurate. SG sometimes catches loop-shadowing bugs the runtime misses — when tightening runtime's loop-variable detection, also revisit `MarkdownValidator`.

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

A substitution token whose dotted path resolves to `IEnumerable<string>` (incl. `string[]`, `List<string>`, `IReadOnlyList<string>`, etc.) auto-renders as a Word bullet list — same output as `{{ Tags | bullet_list }}`. **No attribute required**; detection is type-driven. No SG counterpart, no diagnostics — opportunistic, not validated.

**Runtime path** (`src/Parchment/StringLists/`):

1. `StringListMap.Build(modelType, name)` walks property graph. For each property: skip if `[ExcelsiorTable]` (Excelsior keeps ownership; attribute permits `string` element types and would otherwise be shadowed). Else if assignable to `IEnumerable<string>` (excluding `string` itself, which is `IEnumerable<char>`), add `(DottedPath, Func<object, object?> Getter)`. Same `ChainGetter`/`ShouldDescend`/per-branch `visited` discipline as `ExcelsiorTableMap`.
2. `ScopeTreeRunner.EvaluateTokenAsync` consults `TryResolveStringList` AFTER `TryResolveExcelsiorTable` and `TryResolveFormatted`. Dispatch order: Excelsior → Format → StringList → Fluid.
3. `TryResolveStringList` **gates opportunistically** instead of throwing on misuse:
   - Returns `null` if token isn't solo in its paragraph.
   - Returns `null` if parsed `OutputStatement.Expression` isn't a plain `MemberExpression` (user attached a filter — explicitly opting into Fluid).
   - In both cases, Fluid takes over — preserves backward compat.
4. When gates pass: calls `entry.Getter(rootModel)`, materializes with `.ToList()` (so deferred render delegate doesn't re-enumerate), returns `TokenValueHelpers.BulletList(items)` — same primitive the `bullet_list` filter uses (`TokenValue.OpenXmlToken` wrapping `IOpenXmlContext.CreateBulletNumbering()` + `ListParagraph` paragraphs).
5. `ScopeTreeRunner` takes `StringListMap` as ctor dep, propagates to cloned runners.

**Why opportunistic, not strict**: the existing `bullet_list` filter produces the same output, and pre-feature users wrote `{{ Tags | bullet_list }}` against `IEnumerable<string>` properties — strict validation would have broken `TokenOverrideTests.BulletListFilter`. Fall-through keeps the feature purely additive.

Loop-scoped tokens fall through — same root-only key limitation as `ExcelsiorTableMap`/`FormatMap`. Inside `{% for c in Customers %}{{ c.Tags }}{% endfor %}` use the explicit `bullet_list` filter.

Tests: `StringListTests` (Parchment.Tests/Docx). Scenario: `src/Parchment.Tests/Scenarios/string-list/`.

### Editable field dispatch (`[EditableField]`) + extraction

Two-way binding: marked members render as tagged content controls inside editable-range exceptions, the document is locked read-only, and `ParchmentExtractor` reads user edits back. User-facing docs in `readme.md` → "Editable fields (two-way binding)"; PARCH013–PARCH018 in the diagnostics list.

**Runtime render path** (`src/Parchment/Editable/`):

1. `EditableMap.Build(modelType, name)` mirrors `FormatMap.WalkType` (same `ChainGetter` / `ShouldDescend` / per-branch `visited` discipline). Entries carry getter, **parent-routed setter** (`property.SetValue(parentGetter(root), value)` — never the root; the leaf lives on the object at the end of the path), `CanReach` (parent chain non-null), kind, `IsNullable` (via `NullabilityInfoContext` for refs, `Nullable.GetUnderlyingType` for values), `MultiLine`, `DateFormat`. Model-shape errors throw at build: unsupported type, `bool?`, no public non-init setter (init detected via `IsExternalInit` modreq), conflicting `[ExcelsiorTable]` or `[Markdown]`/`[StringSyntax("markdown")]`. `[Html]`/`[StringSyntax("html")]` is **not** a conflict — it selects the `Html` kind (editable rich text, see the Html sub-point under 4).
2. `EditableTokenValidator` runs after `FormatTokenValidator` in `RegisterDocxTemplate`: plain member access only; duplicate path rejected (body part only — the dotted path is the control tag); editable token sharing a paragraph with an Excelsior/Format token rejected (structural splice/split clones paragraph halves and would corrupt the perm-range markers). A render-time guard in `ScopeTreeRunner` catches the same conflict for `TokenValue`-typed properties that only turn out structural at render.
3. `SettingsProtection.Apply` writes `w:documentProtection w:edit="readOnly" w:enforcement="1"` into the settings part at **registration** (canonical bytes), gated on `!map.IsEmpty && protection == ProtectionMode.WhenEditable`. Passwordless by design (determinism + no real security either way).
4. `ScopeTreeRunner` dispatch order: Excelsior → Format → **Editable** → StringList → Fluid. `TryResolveEditable` returns an internal `EditableToken`; `ProcessSubstitutionAsync` builds the `[permStart, sdt, permEnd]` triple via `EditableFieldBuilder` (sdtPr: site rPr clone, alias, tag = dotted path, sequential id from `EditableState`, `sdtLocked`, kind element; canonical values: `w14:checked`, `w:fullDate`, `w:listItem`) and splices via `EditableSplicer`.
   - **Temporal kinds split.** `DateOnly`/`DateTime` → `Date` kind = native date picker with canonical `w:fullDate` (no display-text parse). `DateTimeOffset` → `DateTimeOffset` kind and `TimeOnly` → `Time` kind, both **plain `w:text` controls** carrying round-trippable ISO run text — because `w:fullDate` is a bare `DateTime` (no offset) and Word has no time-only picker, so the run text is the only faithful carrier. `EditableFieldReader.ReadDateTimeOffset`/`ReadTime` parse it back offset-preserving. `EditableFieldKind.MapKind` (runtime) and `MapEditableKind` (SG) must agree on this split — both list `DateTimeOffset`/`Time` and both live in lockstep with the two `EditableFieldKind` enum copies (runtime `Parchment.Generated`, SG mirror). The old code mapped `DateTimeOffset` onto the date picker and zeroed the offset on extract — regression-guarded by `ExtractTests.TemporalTypesRoundTrip` / `EditedDateTimeOffsetPreservesNewOffset`.
   - **Editable HTML (`Html` kind).** A `string` also marked `[Html]`/`[StringSyntax("html")]` maps to `EditableFieldKind.Html`. `FormatMap.Build` skips it (the editable map claims it — same skip must exist on the SG side). `ProcessSubstitutionAsync` does **not** inline-splice: `QueueEditableHtml` renders the value through `OpenXmlHtml.WordHtmlConverter.ToElements` and `EditableHtmlBuilder` wraps the produced blocks in a **block-level** `[permStart, SdtBlock(sdtLocked, tag = path), permEnd]` triple (the block analogue of `EditableFieldBuilder`; a rich-text control has no kind element). It must sit solo in its paragraph (own-paragraph rule, like read-only `[Html]`) — enforced at render, not registration. `[Markdown]` + editable stays rejected — extraction has no OpenXML→Markdown serializer.
5. **Body-only**: `RegisteredDocxTemplate.RenderPartAsync` passes `EditableMap.Empty` for non-body parts — header/footer occurrences render as plain read-only mirrors. `ProcessLoopAsync`'s cloned runner also gets `EditableMap.Empty` (per-iteration controls would duplicate tags); if-branch inner runners keep the real map.

**Extraction** (`src/Parchment/Extraction/`): `ParchmentExtractor.Extract<T>` opens read-only, matches `body.Descendants<SdtElement>()` tags against the map (per-type cached — Extract is per-request), reads per kind via `EditableFieldReader`, reports `FieldState` per field (`Extracted`/`Empty`/`Missing`/`ParseFailed`/`Duplicate` — first occurrence wins). `ExtractResult.ApplyTo` validates `CanReach` for every applicable field **before** mutating anything, then assigns; `Empty` → null for nullable members only. Default culture is `CultureInfo.InvariantCulture` **because** Fluid's `TemplateOptions.CultureInfo` defaults to invariant and `SharedFluid` doesn't override it — the render and parse cultures must match, and both default the same way. The `Html` kind reads via `OpenXmlHtmlSerializer.ToHtml(sdtContent, mainPart)` — the reverse of `WordHtmlConverter`, walking **standard** OOXML (pStyle headings, `w:numPr` lists resolved through the numbering part, run `rPr` emphasis, hyperlink rels) so it inverts both the render and Word's own edits; subset = `p`/`h1`–`h6`/`strong`/`em`/`u`/`s`/`sup`/`sub`/`a`/`br`/`ul`/`ol`, anything else degrades to text. Because it needs the numbering/hyperlink parts, `EditableFieldReader.Read` takes `MainDocumentPart`.

**SG path** (`src/Parchment.SourceGenerator/`):

1. `ShapeBuilder` detects `Parchment.EditableFieldAttribute` via typed symbol (`GetTypeByMetadataName`, like Excelsior), reads `MultiLine`/`DateFormat` named args, maps the kind from the symbol (`EditableKindFor`/`MapEditableKind` — lockstep with runtime `EditableMap`; `[Html]`+string → `Html` kind, and `AccessorEmission` keeps such members out of `_Formats`, mirroring `FormatMap.Build`'s editable skip), and bakes primitives into `MemberEntry` (`IsEditable`, `EditableKind?`, `EditableIsNullable`, `HasUsableSetter`, `EditableMultiLine`, `EditableDateFormat`).
2. `ValidateEditableShape` fires PARCH013/014/015 per offending member whether or not a token references it — lockstep with `EditableMap.Build` throwing unconditionally at docx registration. PARCH015 now fires only for `[ExcelsiorTable]`/`[Markdown]` + editable; `[Html]` + editable is valid. Docx targets only; the markdown flow never builds the map, so markdown targets stay silent.
3. Token rules run in a **separate body-scoped pass** (`ValidateEditableTokens`): `DocxArchiveReader`/`DocxData` carry `BodyParagraphs` alongside the all-parts list (zip entry order isn't body-first, so the body part is tracked explicitly). The pass tracks loop scope silently (no PARCH001/005 re-reports) and emits PARCH016 (non-plain), PARCH017 (duplicate in body — header mirrors don't count), PARCH018 warning (any editable-resolved token inside a `{% for %}` body, root-pathed or loop-var-pathed).
4. `AccessorEmission` emits the fifth dataset `_Editables` (`EditableFieldMapEntry`: path, kind, `typeof(underlying)`, nullability, getter, block-bodied parent-routed setter with nullability-aware cast — `(decimal?)v` vs `(string)v!` — and `CanReach` lambda) registered via `GeneratedRegistration.RegisterEditable`. Invalid members (kind null / no setter) are skipped so emitted source stays compilable; the error diagnostic fails the build anyway.
5. `[ParchmentModel(Protection = ...)]` arrives as the enum's underlying int (SG can't reference Parchment.dll), lands on `TargetInfo` as the SG-side `ProtectionMode`, and is emitted as a third `RegisterDocxTemplate` argument **only when non-default** — keeping pre-existing generated output (and snapshots) byte-identical.

**Lockstep**: runtime `EditableMap` walk/rules ↔ `ShapeBuilder`+`ValidateEditableShape`; runtime `EditableTokenValidator` ↔ `ValidateEditableTokens`; runtime `EditableMap.BuildSetter` parent-routing ↔ `AccessorEmission.EmitSetter`. Tighten in the same PR. Tests: `EditableFieldTests` + `ExtractTests` (runtime), `EditableFieldGeneratorTests` (SG). Scenario: `src/Parchment.Tests/Scenarios/editable-fields/`.

### Determinism guarantee

See `readme.md` → "Determinism". Implementation discipline: avoid `w:rsid` randomness, never set `PackageProperties.Created`, no timestamps. Editable fields: sequential sdt/perm-range ids (`EditableState`), passwordless `w:documentProtection` (a password would need a random salt). `DeterminismTests.cs` renders samples twice (including scalar and HTML editable-field cases) and asserts byte equality — don't break it. Caveat: a **pre-existing OpenXmlHtml** gap is out of scope here — rendering `[Html]` with a list into a template that lacks a numbering part draws a non-deterministic numbering relationship id, so `HtmlEditableFieldRenderIsByteIdentical` deliberately excludes lists.

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
    └── output#00.verified.txt        # Verify text extraction
```

**Test wiring** (see `ExcelsiorTableTests.Render`):

1. `.cs` file lives under `src/Parchment.Tests/Docx/` (or wherever feature tests live), not inside the scenario dir — scenario dir is asset-only.
2. Reads `input.docx` via a `[CallerFilePath]`-anchored helper:
   ```csharp
   static string SourcePath([CallerFilePath] string path = "") => path;
   static string ScenarioPath(string name) =>
       Path.GetFullPath(Path.Combine(Path.GetDirectoryName(SourcePath()) ?? "", "..", "Scenarios", name));
   ```
3. Directs Verify into the scenario dir with a custom prefix: `settings.UseDirectory(ScenarioPath("excelsior-table"))` + `settings.UseFileName("output")` → `output.verified.docx` rather than `ClassName.MethodName.*`.

**Regenerating `input.png`**: `ScenarioInputRenderer.cs` is `[Test, Explicit]` — excluded from default runs. It globs `Scenarios/**/input.docx`, renders each via Morph/SkiaSharp, writes the first page next to the source. Invoke on demand:

```bash
dotnet run --project src/Parchment.Tests --configuration Release -- \
    --treenode-filter "/*/*/ScenarioInputRenderer/RenderAllInputDocxesToPng"
```

`input.png` is *committed* alongside `input.docx`.

**Adding a scenario**: mkdir under `Scenarios/`, drop `input.docx` (binary already in `.gitattributes`), write a feature test that calls `UseDirectory(ScenarioPath("<name>")) + UseFileName("output")`, run `ScenarioInputRenderer` once for `input.png`, reference both PNGs from `readme.md` (note `#` → `%23` URL escape).

## Design decisions

### Source generator emits pre-compiled accessors; runtime reflection is a fallback

User-facing summary in `readme.md` → "Source generator (recommended)". The SG path is the **primary** registration path; the runtime reflection walks in `SharedFluid.RegisterTypeGraph`, `ExcelsiorTableMap.WalkType`, `FormatMap.WalkType`, and `StringListMap.WalkType` are kept only as a fallback for callers that can't use the SG. Pattern mirrors `System.Text.Json`'s `JsonSerializerContext`, EF Core's compiled models, and the regex source generator.

**Pipeline.** `AccessorEmission.Emit(target.Shape, rootFqn)` walks `ModelShape` and produces four parallel datasets:

1. **Fluid accessors**: `KeyValuePair<string, IMemberAccessor>[]` per non-empty type (one `DelegateAccessor` per public property/field), registered via `GeneratedRegistration.RegisterFluidAccessors(typeof(T), arr)` → `SharedFluid.registeredTypes` + `Options.MemberAccessStrategy.Register`.
2. **Excelsior-table entries**: dotted-path walk collecting `[ExcelsiorTable]` members as `(DottedPath, ElementType, Getter)`. Registered via `RegisterExcelsiorTable` → `ExcelsiorTableMap.precompiledCache`.
3. **Format entries**: same walk for `[Html]` / `[Markdown]` / `[StringSyntax]` strings as `(DottedPath, FormatMapKind, Getter)`. Registration walk AND render-time `ScopeTreeRunner.TryResolveFormatted` use the emitted getter — no per-segment reflection on SG path.
4. **String-list entries**: dotted-path walk for `IEnumerable<string>` members not already owned by `[ExcelsiorTable]`. Populates `StringListMap.precompiledCache`.

All four runtime maps have `static readonly ConcurrentDictionary<Type, X> precompiledCache` checked at the top of `Build(modelType, ...)`. Cache hit returns the SG-built map and skips the reflection walk entirely; miss falls through to reflection (runtime API contract unchanged).

**Public API surface.** SG-emitted code calls into `Parchment.Generated`:
- `GeneratedRegistration` — static class with one `Register*` method per dataset.
- `ExcelsiorTableMapEntry`, `FormatMapEntry`, `StringListMapEntry` — public data carrier records.
- `FormatMapKind` — public enum (replaces internal `FormatKind`).

These must stay public — SG runs against user assemblies, signed assembly + arbitrary user PKs means `[InternalsVisibleTo]` won't help. The maps themselves (`SharedFluid`, `ExcelsiorTableMap`, etc.) stay internal.

**Path-getter emission.** Cast-and-null-chain lambdas: `o => ((global::NS.Root)o).First?.Second?.Third`. Root cast unconditional (`Render` rejects null models); subsequent segments use `?.` to match runtime `ChainGetter` null-short-circuit. Non-nullable value-type intermediates would fail to compile — accepted, since binding-model intermediates are virtually always reference types.

**Lockstep.** `AccessorEmission.WalkForMaps` (SG, Roslyn symbols + flags baked into `MemberEntry`) and runtime `WalkType` implementations must produce equivalent results: same dotted-path key set, element types, kinds. When tightening one, revisit the other. Tests: `Parchment.Tests` (runtime), `Parchment.SourceGenerator.Tests` (SG snapshots), `IntegrationTests` (SG output against packed nupkg).

### `[ParchmentModel]` lives on the binding model, not on an intermediary "template" class

Full rationale and alternatives in `readme.md` → "Source generator (recommended)". Implementation-relevant points:

- The attribute target IS the model; no separate marker class, no `typeof(TModel)` argument.
- Model must be `partial` (SG emits `RegisterWith` onto it). PARCH011 fires when nested inside a non-partial enclosing type.
- `[AttributeUsage(AttributeTargets.Class)]` blocks interface targets (CS0592). Runtime interface guard: `GuardBindingModel<TModel>` in `TemplateStore.cs`.
- One template per model via the attribute is canonical. Multi-template / POCO / dynamic-path scenarios use the runtime `TemplateStore.RegisterDocxTemplate<T>` / `RegisterMarkdownTemplate<T>` overloads — keep them as the supported escape hatch.
- When touching `ParchmentTemplateGenerator.cs`, the attribute predicate, target shape, `BuildPartialSource` wrapping, and any new diagnostics all treat the attribute target as the model type. No second symbol to thread through.

## Non-obvious gotchas

- **Tokens straddling run boundaries**: Word splits text into multiple `<w:r>` when formatting changes, proofing markers fire, or smart-quote autocorrect runs. `{{ customer.name }}` can land across N runs. Scanner uses `paragraph.InnerText` + `RunMap` (offset → `<w:t>`) so substitutions land correctly. Formatting of the **first run** containing the opening `{{` wins for the entire substitution.

- **PascalCase tokens**: Parchment liquid uses PascalCase (`{{ Customer.Name }}`); Fluid's default member access is case-insensitive. No snake-case translation — an earlier `MemberNameStrategies.SnakeCase` attempt was abandoned (API doesn't exist in Fluid 2.15).

- **Method-call syntax (`{{ Customer.FullName() }}`) fails at registration, not silently at render**: Fluid's `AllowFunctions` stays off — Parchment's binding model is data, not behaviour. `FluidParser.TryParse` returns false; runtime surfaces `ParchmentRegistrationException`, SG surfaces `PARCH006` (markdown) / parse diagnostic (docx). Workaround: expose a computed property.

- **Indexer syntax with a string literal (`{{ Customer['FullName'] }}`) is validated the same as dotted access**: Fluid treats both as the same member resolution at render time. The three path collectors (`Liquid/IdentifierVisitor.cs`, `Parchment.SourceGenerator/IdentifierVisitor.cs`, `MarkdownValidator.ExpressionPathCollector`) extract the literal from `IndexerSegment { Expression: LiteralExpression { Value.Type: FluidValues.String } }` and treat it as `IdentifierSegment`. Non-literal indexers (numeric, variable, expression) still terminate the path walk — those are dictionary/array access at runtime, not statically validatable. Shared extractor: `Liquid/IdentifierVisitor.TryGetStaticName` (runtime) and `Parchment.SourceGenerator/SegmentNames.TryGetStaticName` (SG); mirror when tightening. Side effect: Excelsior / Format / StringList dispatch also accepts indexer notation. Tests: `ReferenceValidatorTests.Indexer*`, `ParchmentTemplateGeneratorTests.*IndexerStringLiteral*`.

- **`ScopeTreeRunner.ProcessLoopAsync` attaches each iteration's clones to a scratch `Body` before running the nested scope tree**. Without this, nested `{% for %}`/`{% if %}` silently no-op: `open.Parent` returns null on a detached clone, `CaptureBetween` captures nothing, the inner block-tag text lands as literal `{% for ... %}`. Reverting to `parent.InsertAfter(clone, insertAnchor)` *before* the nested run breaks nested loops in a way only `LoopTests.NestedLoop` catches.

- **`ScopeTreeRunner.ProcessIfAsync` keeps the chosen branch positionally, not via anchors**: static paragraphs and tables inside a branch have no anchor bookmarks and no scope-tree nodes (`TokenScanner.Scan` skips token-free paragraphs), so the runner captures the chosen branch's physical range — everything between its tag paragraph (`IfBranch.TagAnchorName` / `IfNode.ElseAnchorName`) and the next boundary tag (`{% elsif %}` / `{% else %}` / `{% endif %}`). An anchor-derived keep-set silently drops static/table content in the chosen branch; `ConditionalTests.StaticParagraphInChosenBranchIsKept` and `TableInsideChosenBranchIsKeptAndSubstituted` guard against regressing. Same reason `IfNode.ElseAnchorName` (not `ElseBody.Count`) signals an else branch: a static-only else has an empty body but still owns a physical range.

- **`SettingsProtection` must respect the CT_Settings sequence**: `w:settings` is a strict xsd:sequence — appending `w:documentProtection` at the end produces a file Word "repairs". Insertion goes before the first child not in the preceding-types set. Two naming traps when touching that set: the settings-level `w:trackChanges` element is SDK class `TrackRevisions` (NOT `TrackChanges`, which doesn't exist), and `RemoveDateAndTime` sits between `RemovePersonalInformation` and `DoNotDisplayPageBoundaries`. `EditableFieldTests.OutputValidates` runs `OpenXmlValidator` over the output — it catches ordering mistakes.

- **`EditableSplicer` exists because `ParagraphSplicer` duplicates zero-width elements**: the clone-and-trim approach distributes elements with no text (bookmarks, `permStart`/`permEnd`, emptied sdt shells) into BOTH halves — harmless for anchor bookmarks (StripAll removes them), fatal for editable fields (duplicate perm ids, stray control shells) when a paragraph hosts more than one. `EditableSplicer` splits runs at exact token boundaries instead. Don't "simplify" editable splicing back onto `ParagraphSplicer.SpliceInline`; `EditableFieldTests.TwoEditableFieldsInOneParagraph` and `TokenStraddlingRunsSplicesCorrectly` guard it.

- **SG-side `EditableFieldKind` / `ProtectionMode` are copies matched by member name**: the generator can't reference Parchment.dll, so emission writes `global::Parchment.Generated.EditableFieldKind.{name}` / `global::Parchment.ProtectionMode.{name}` from its own enum's `ToString()`. Renaming or reordering members on either side silently breaks generated code or protection pass-through — change both in the same PR (same class of trap as the `[ExcelsiorTable]` FQN-string match).

- **`OpenXmlMarkdownRenderer` is not thread-safe** — one instance per render. The `Stack<ContainerState>` and `ObjectRenderers` are mutable. The `RegisteredTemplate` (cached canonical bytes + scope tree) IS immutable — concurrent renders work, each gets its own renderer.

- **`appveyor.yml` font validation** — every TTF/OTF in `src/Fonts/` is loaded through `PrivateFontCollection` BEFORE copy to `%WINDIR%\Fonts`, catching Git CRLF corruption upfront. When adding a font, mark it binary in `.gitattributes` (already present for `*.ttf`, `*.otf`).

- **`ParchmentModel` is a separate project** (not `Model`) to avoid clashes with common test fixture names in IDE autocomplete.

- **Excelsior dispatch bypasses Fluid, deliberately**: `ExcelsiorTableBridge` walks the CLR model directly via cached `Func<object, object?>` getter chains. Routing through Fluid would erase `IEnumerable<T>` (`ArrayValue.ToObjectValue()` returns `FluidValue[]`) — `new WordTableBuilder<T>(data)` needs the typed enumerable. This is why `ScopeTreeRunner` takes `rootModel` as a separate ctor param. Removing it breaks Excelsior on the first filtered or loop-nested token.

- **Parchment is binary-coupled to Excelsior's `WordTableBuilder` constructor**: `ExcelsiorTableBridge.BuildTable<T>` calls `new WordTableBuilder<T>(data).Build(mainPart)` in *compiled* IL (only the generic-method dispatch is reflection). So an Excelsior release that changes `WordTableBuilder`'s ctor signature — **even adding an optional parameter** (which the compiler bakes into the call site, changing the binary arity) — breaks the already-compiled Parchment: consumers hit `MissingMethodException: WordTableBuilder..ctor(...)` at render time. Bumping the `Excelsior` `PackageReference` is **not** enough — Parchment must be rebuilt and republished against the new Excelsior. Conversely, additive Excelsior changes (new fluent methods like `BodyStyle`/`BodyParagraphStyle`) are safe and need no Parchment rebuild. Symptom seen in the wild: every `[ExcelsiorTable]` report throwing while loop-built tables (which don't go through the bridge) render fine.

- **Per-branch visited set in `ExcelsiorTableMap.WalkType`**: cycle prevention uses `HashSet<Type>` with `Add` on descend / `Remove` on return. Same type at unrelated paths (e.g. `Order.Buyer.Addresses` + `Order.Seller.Addresses`) still visits both; a self-reference (`Node.Next` → `Node`) is pruned. A global no-removal visited set silently drops the second sibling branch.

- **SG `[ExcelsiorTable]` detection matches by FQN string** (`"global::Parchment.ExcelsiorTableAttribute"` in `ShapeBuilder.HasExcelsiorTableAttribute`) — SG can't `typeof()` it. Renaming/moving the attribute silently breaks `PARCH007`/`PARCH008` until the literal is updated.

- **Excelsior runtime and SG validators must stay in lockstep**: `ExcelsiorTokenValidator` (runtime) and `ValidateExcelsiorToken` + `ShapeResolver.IsExcelsiorTableMember` (SG) enforce the same two rules — solo-in-paragraph and plain-member-access. Runtime checks Fluid AST directly; SG piggybacks on `Token.IsPlainIdentifier` from `TokenScanner.IsPlainMemberAccess`. Tighten both in the same PR.

- **`MemberEntry.IsExcelsiorTable` must stay primitive** — it's a `bool` on a `sealed record` flowing through the incremental generator pipeline; equality is structural. Adding a `bool` was safe; adding `List<T>`, an `ISymbol` reference, or any mutable field defeats cacheability and forces ShapeBuilder to re-run on every compilation.

- **Package dep direction is Parchment → Excelsior, hard**. The `[ExcelsiorTable]` attribute deliberately lives in Parchment, not Excelsior. Moving it to Excelsior would invert the dep: every Excel-only Excelsior consumer would pull Parchment, and every Parchment user would have to reference an attribute-only stub package. Don't.

- **SG markdown reads via `AdditionalText.GetText`, not `File.ReadAllText`**: RS1035 bans `File.*` in analyzers. `ParchmentTemplateGenerator.ReadMarkdown` calls `text.GetText(cancel)` and treats null as a read error. Test harness's `PathAdditionalText.GetText` returns `SourceText` for `.md`/`.markdown` but `null` for `.docx` (docx branch uses `ZipFile.OpenRead(text.Path)`). Failure mode is `PARCH006` "AdditionalText returned no SourceText" — loud, not silent.

- **Public fields are bindable at every depth, and the runtime registration walks them explicitly**: `SharedFluid.RegisterTypeGraph` registers properties via Fluid's `MemberAccessStrategyExtensions.Register<T>(strategy)` (properties only), THEN enumerates `GetFields(BindingFlags.Public | BindingFlags.Instance)` and registers each as a `DelegateAccessor`. Both properties' AND fields' types are recursed. Without the field-walk pass, top-level field access still works (`TemplateContext(model, ..., allowModelMembers: true)` reflects on the root instance) but nested-via-field access (`{{ Outer.FieldOnInner }}`) renders empty. SG-side `ShapeBuilder.TryGetMemberType` and `ModelValidator.ResolveMember` already handled fields — the Fluid registration step was the only gap.

- **Model binding limitations** — see `readme.md` → "Model binding limitations" for the user-facing list (interfaces as `TModel`, interface default impls, C# 14 extension properties, explicit interface implementations all unsupported). Implementation: runtime interface block is `GuardBindingModel<TModel>` in `TemplateStore.cs` (abstract classes supported; tests in `InterfaceModelTests`). SG block is `[AttributeUsage(AttributeTargets.Class)]` on `ParchmentModelAttribute` (CS0592). For the other three, reflection walkers and `ShapeBuilder.BuildEntry` deliberately don't walk `GetInterfaces()` / `current.Interfaces` and never visit extension containers. Inheritance from a base **class** IS supported (`ReferenceValidatorTests.InheritedMember_*`, `ParchmentTemplateGeneratorTests.InheritedMember_Validates`).

- **`MarkdownValidator` binds the loop variable to the root type when source is unresolved**: in `WalkFor`, a loop whose `Source` fails (PARCH001/002) still binds `Identifier` → `target.Shape.RootTypeFullyQualifiedName` for the body walk. Intentional cascade-suppression — stops every body access (`{{ line.Description }}` etc.) from also tripping PARCH001. Skipping the body walk would stop validating nested constructs; binding to a sentinel would generate false positives. Wrong-but-optimal — keep it.
