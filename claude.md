# CLAUDE.md

- User-facing feature docs, PARCH diagnostic codes, source-generator usage, model binding limitations, and the determinism guarantee live in `readme.md`.
- Build/test commands, architecture, design decisions, and non-obvious gotchas live in `contributing.md`.

Read both before making changes. When updating either, keep the cross-references intact rather than duplicating content here.

## Running tests

Tests use **TUnit**, not VSTest. `dotnet test` is unsupported on .NET 10 SDK and will error. Use `dotnet run` against the test project, and TUnit's `--treenode-filter` (not `--filter`) for narrowing:

```bash
# All tests in the main suite
dotnet run --project src/Parchment.Tests --configuration Release

# Single class
dotnet run --project src/Parchment.Tests --configuration Release -- --treenode-filter "/*/*/HtmlInlineRendererTests/*"

# Single test
dotnet run --project src/Parchment.Tests --configuration Release -- --treenode-filter "/*/*/HtmlInlineRendererTests/ITagAppliesItalic"
```

Other test projects: `src/Parchment.SourceGenerator.Tests`, `IntegrationTests/IntegrationTests` (the latter requires `src` to be packed first).
