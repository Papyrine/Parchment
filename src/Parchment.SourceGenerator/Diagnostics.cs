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
        title: "Enclosing type of [ParchmentModel] target must be partial",
        messageFormat: "Model '{0}' is nested inside '{1}' which is not declared partial; the source generator emits the registration helper as a partial declaration and every enclosing type on the chain must be partial too",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor MissingRemovePersonalInformation = new(
        id: "PARCH012",
        title: "Template missing 'Remove personal information on save' setting",
        messageFormat: "Template '{0}' does not have the Word 'Remove personal information from file properties on save' setting enabled. Enable it via File → Options → Trust Center → Trust Center Settings → Privacy Options.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EditableUnsupportedType = new(
        id: "PARCH013",
        title: "[EditableField] member has an unsupported type",
        messageFormat: "Model '{0}' member '{1}' is [EditableField] but its type '{2}' is not supported. Supported: string, bool, DateOnly, DateTime, DateTimeOffset, TimeOnly, enums, and numeric types (nullable variants except bool? — a checkbox cannot represent null).",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EditableNoSetter = new(
        id: "PARCH014",
        title: "[EditableField] member has no usable setter",
        messageFormat: "Model '{0}' member '{1}' is [EditableField] but has no public non-init setter; extraction writes values back onto the model, so the member must be settable",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EditableConflictingAttribute = new(
        id: "PARCH015",
        title: "[EditableField] combined with a conflicting attribute",
        messageFormat: "Model '{0}' member '{1}' combines [EditableField] with [ExcelsiorTable] or [Markdown]; editable rich text is supported via [Html] only, and other formats are plain typed content, not rendered markup",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EditableTokenNotPlainIdentifier = new(
        id: "PARCH016",
        title: "[EditableField] token must be a plain member-access expression",
        messageFormat: "Template '{0}' token '{1}' references an [EditableField] member with filters or a non-plain expression; the editable render path is selected by attribute so filters would not be applied",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EditableTokenDuplicated = new(
        id: "PARCH017",
        title: "[EditableField] member referenced more than once in the document body",
        messageFormat: "Template '{0}' token '{1}' references an [EditableField] member already referenced elsewhere in the document body; the dotted path is the content control's tag and must be unique for extraction",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor EditableTokenInLoop = new(
        id: "PARCH018",
        title: "[EditableField] token inside a loop renders read-only",
        messageFormat: "Template '{0}' token '{1}' references an [EditableField] member inside a '{{% for %}}' body; loop iterations would produce duplicate control tags, so the token renders as plain read-only text instead of an editable field",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor RenderAttributeOnStaticMember = new(
        id: "PARCH019",
        title: "Render attribute on a static member has no effect",
        messageFormat: "Member '{0}' is static, so '[{1}]' has no effect; the per-template maps that dispatch it walk instance members only. The value still binds and renders as plain text. Make the member non-static, or drop the attribute.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);
}
