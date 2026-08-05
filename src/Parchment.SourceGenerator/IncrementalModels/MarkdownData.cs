sealed record MarkdownData(
    string Path,
    string Text,
    string? ReadError,
    // Set when the template is declared with <ParchmentEmbeddedTemplate>: the manifest name its
    // staged copy is embedded under, supplied by Parchment.targets so the generated registration
    // and the embedding cannot disagree. Null for a template that lands on disk instead.
    string? ResourceName = null);
