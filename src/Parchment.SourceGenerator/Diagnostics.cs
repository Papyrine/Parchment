static class Diagnostics
{
    public static readonly DiagnosticDescriptor MissingMember = new(
        id: "PARCH001",
        title: "Template references an unknown model member",
        messageFormat: "Template '{0}' token '{1}' references '{2}' which is not a member of '{3}'",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor LoopSourceNotEnumerable = new(
        id: "PARCH002",
        title: "Loop source is not enumerable",
        messageFormat: "Template '{0}' loop '{1}' source does not resolve to a type implementing IEnumerable<T>",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnsupportedBlockTag = new(
        id: "PARCH003",
        title: "Unsupported block tag",
        messageFormat: "Template '{0}' uses unsupported block tag '{1}' (supported: for, endfor, if, elsif, else, endif)",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor TemplateFileMissing = new(
        id: "PARCH004",
        title: "Template file not found in AdditionalFiles",
        messageFormat: "Template path '{0}' was not found in AdditionalFiles — add <AdditionalFiles Include=\"...\"/> to the csproj",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor MixedInlineBlockTag = new(
        id: "PARCH005",
        title: "Block tag must sit in its own paragraph",
        messageFormat: "Template '{0}' block tag '{1}' shares a paragraph with other content; block tags must be on their own lines",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor TemplateReadError = new(
        id: "PARCH006",
        title: "Failed to read template",
        messageFormat: "Template '{0}' could not be read: {1}",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor ExcelsiorTokenNotAlone = new(
        id: "PARCH007",
        title: "[ExcelsiorTable] token must sit alone in its own paragraph",
        messageFormat: "Template '{0}' token '{1}' references an [ExcelsiorTable] property but shares its paragraph with other content; structural table replacement would discard the surrounding text",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    // PARCH009 was previously emitted for `[Html]`/`[Markdown]` tokens that did not sit alone in
    // their paragraph. The runtime now splices inline content in place and splits the host
    // paragraph for block-level content, so non-solo tokens are valid. The id is intentionally
    // not reused.

    public static readonly DiagnosticDescriptor FormatTokenNotPlainIdentifier = new(
        id: "PARCH010",
        title: "[Html]/[Markdown] token must be a plain member-access expression",
        messageFormat: "Template '{0}' token '{1}' references an [Html]/[Markdown] property with filters or a non-plain expression; the property's formatted rendering is selected by attribute so filters would not be applied",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor ExcelsiorTokenNotPlainIdentifier = new(
        id: "PARCH008",
        title: "[ExcelsiorTable] token must be a plain member-access expression",
        messageFormat: "Template '{0}' token '{1}' references an [ExcelsiorTable] property with filters or a non-plain expression; the Excelsior render path bypasses Fluid and walks the model directly, so filters would be ignored",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EnclosingTypeNotPartial = new(
        id: "PARCH011",
        title: "Enclosing type of [ParchmentTemplate] target must be partial",
        messageFormat: "Template '{0}' is nested inside '{1}' which is not declared partial; the source generator emits the registration helper as a partial declaration and every enclosing type on the chain must be partial too",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

}
