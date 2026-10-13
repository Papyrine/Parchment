# CLAUDE.md

- User-facing feature docs, PARCH diagnostic codes, source-generator usage, model binding limitations, and the determinism guarantee live in `readme.md`.
- Build/test commands, architecture, design decisions, and non-obvious gotchas live in `contributing.md`.

Read both before making changes. When updating either, keep the cross-references intact rather than duplicating content here.

## Running tests

Tests use **TUnit**, not VSTest. `dotnet test` is unsupported on .NET 10 SDK and will error. Use `dotnet run` against the test project, and TUnit's `--treenode-filter` (not `--filter`) for narrowing.

Local `Release` builds also need `-p:IsPackable=false`, or SponsorCheck fails the build with `SC100: Platform fetch failed. GitHub GraphQL HTTP 401` — it only bundles on CI, where it has credentials. This applies to every `build`, `run` and `test` invocation below, and to the sibling repos too.

```bash
# All tests in the main suite
dotnet run --project src/Parchment.Tests --configuration Release -p:IsPackable=false

# Single class
dotnet run --project src/Parchment.Tests --configuration Release -p:IsPackable=false -- --treenode-filter "/*/*/HtmlInlineRendererTests/*"

# Single test
dotnet run --project src/Parchment.Tests --configuration Release -p:IsPackable=false -- --treenode-filter "/*/*/HtmlInlineRendererTests/ITagAppliesItalic"
```

Other test projects: `src/Parchment.SourceGenerator.Tests`, `IntegrationTests/IntegrationTests` (the latter requires `src` to be packed first).

Work often spans the sibling repos `../OpenXmlHtml` and `../Excelsior`. Both also run **TUnit** on Microsoft.Testing.Platform, so `--filter` matches nothing there either: narrow with `--treenode-filter`. Unlike here, their `global.json` selects the Microsoft.Testing.Platform runner, so `dotnet test` works (`dotnet test OpenXmlHtml.slnx`, `dotnet test --project src/Excelsior.Tests`). Each has its own `CLAUDE.md`; read it for the rest, but its `--filter` examples are stale.

## Scan log

- 2026-10-11: perf and bug scan.
