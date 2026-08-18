static class Diagnostics
{
    public static readonly DiagnosticDescriptor MissingMember = new(
        id: "PARCH001",
        title: "Template references an unknown model member",
        messageFormat: "Template '{0}' token '{1}' references '{2}' which is not a member of '{3}'",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch001--unknown-model-member");

    public static readonly DiagnosticDescriptor LoopSourceNotEnumerable = new(
        id: "PARCH002",
        title: "Loop source is not enumerable",
        messageFormat: "Template '{0}' loop '{1}' source does not resolve to a type implementing IEnumerable<T>",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch002--loop-source-is-not-enumerable");

    public static readonly DiagnosticDescriptor UnsupportedBlockTag = new(
        id: "PARCH003",
        title: "Unsupported block tag",
        messageFormat: "Template '{0}' uses unsupported block tag '{1}' (supported: for, endfor, if, elsif, else, endif)",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch003--unsupported-block-tag");

    public static readonly DiagnosticDescriptor TemplateFileMissing = new(
        id: "PARCH004",
        title: "No template found for model",
        messageFormat: "Model '{0}' has no template: nothing is named '{0}.parchment.docx' or '{0}.parchment.md'. {1} A file " +
                       "with that name is picked up wherever it sits in the project; one outside the project needs " +
                       "<ParchmentTemplate Include=\"...\"/> in the csproj. If the file is there and only the IDE reports " +
                       "this, it likely carries a second item type that hides it from the IDE's generator host.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch004--no-template-found-for-model");

    public static readonly DiagnosticDescriptor MixedInlineBlockTag = new(
        id: "PARCH005",
        title: "Block tag must sit in its own paragraph",
        messageFormat: "Template '{0}' block tag '{1}' shares a paragraph with other content; block tags must be on their own lines",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch005--block-tag-shares-a-paragraph");

    public static readonly DiagnosticDescriptor TemplateReadError = new(
        id: "PARCH006",
        title: "Failed to read template",
        messageFormat: "Template '{0}' could not be read: {1}",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch006--template-file-unreadable");

    public static readonly DiagnosticDescriptor ExcelsiorTokenNotAlone = new(
        id: "PARCH007",
        title: "[ExcelsiorTable] token must sit alone in its own block",
        messageFormat: "Template '{0}' token '{1}' references an [ExcelsiorTable] property but shares its block with other content — its paragraph in a docx template, its line in a markdown one; the table replaces the whole block, so the surrounding text would be discarded",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch007--excelsiortable-token-not-alone-in-its-block");

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
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch010--html--markdown-token-with-filters-or-complex-expression");

    public static readonly DiagnosticDescriptor ExcelsiorTokenNotPlainIdentifier = new(
        id: "PARCH008",
        title: "[ExcelsiorTable] token must be a plain member-access expression",
        messageFormat: "Template '{0}' token '{1}' references an [ExcelsiorTable] property with filters or a non-plain expression; the Excelsior render path bypasses Fluid and walks the model directly, so filters would be ignored",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch008--excelsiortable-token-with-filters-or-complex-expression");

    public static readonly DiagnosticDescriptor EnclosingTypeNotPartial = new(
        id: "PARCH011",
        title: "Enclosing type of [ParchmentModel] target must be partial",
        messageFormat: "Model '{0}' is nested inside '{1}' which is not declared partial; the source generator emits the registration helper as a partial declaration and every enclosing type on the chain must be partial too",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch011--enclosing-type-of-parchmentmodel-target-must-be-partial");

    public static readonly DiagnosticDescriptor MissingRemovePersonalInformation = new(
        id: "PARCH012",
        title: "Template missing 'Remove personal information on save' setting",
        messageFormat: "Template '{0}' does not have the Word 'Remove personal information from file properties on save' setting enabled. Enable it via File → Options → Trust Center → Trust Center Settings → Privacy Options.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch012--template-missing-remove-personal-information-on-save-setting");

    public static readonly DiagnosticDescriptor EditableUnsupportedType = new(
        id: "PARCH013",
        title: "[EditableField] member has an unsupported type",
        messageFormat: "Model '{0}' member '{1}' is [EditableField] but its type '{2}' is not supported. Supported: string, bool, DateOnly, DateTime, DateTimeOffset, TimeOnly, enums, and numeric types (nullable variants except bool? — a checkbox cannot represent null).",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch013--editablefield-member-has-an-unsupported-type");

    public static readonly DiagnosticDescriptor EditableNoSetter = new(
        id: "PARCH014",
        title: "[EditableField] member has no usable setter",
        messageFormat: "Model '{0}' member '{1}' is [EditableField] but has no public non-init setter; extraction writes values back onto the model, so the member must be settable",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch014--editablefield-member-has-no-usable-setter");

    public static readonly DiagnosticDescriptor EditableConflictingAttribute = new(
        id: "PARCH015",
        title: "[EditableField] combined with a conflicting attribute",
        messageFormat: "Model '{0}' member '{1}' combines [EditableField] with [ExcelsiorTable] or [Markdown]; editable rich text is supported via [Html] only, and other formats are plain typed content, not rendered markup",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch015--editablefield-combined-with-a-conflicting-attribute");

    public static readonly DiagnosticDescriptor EditableTokenNotPlainIdentifier = new(
        id: "PARCH016",
        title: "[EditableField] token must be a plain member-access expression",
        messageFormat: "Template '{0}' token '{1}' references an [EditableField] member with filters or a non-plain expression; the editable render path is selected by attribute so filters would not be applied",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch016--editablefield-token-with-filters-or-complex-expression");

    public static readonly DiagnosticDescriptor EditableTokenDuplicated = new(
        id: "PARCH017",
        title: "[EditableField] member referenced more than once in the document body",
        messageFormat: "Template '{0}' token '{1}' references an [EditableField] member already referenced elsewhere in the document body; the dotted path is the content control's tag and must be unique for extraction",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch017--editablefield-member-referenced-more-than-once-in-the-body");

    public static readonly DiagnosticDescriptor EditableTokenInLoop = new(
        id: "PARCH018",
        title: "[EditableField] token inside a loop renders read-only",
        messageFormat: "Template '{0}' token '{1}' references an [EditableField] member inside a '{{% for %}}' body; loop iterations would produce duplicate control tags, so the token renders as plain read-only text instead of an editable field",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch018--editablefield-token-inside-a-loop");

    public static readonly DiagnosticDescriptor RenderAttributeOnStaticMember = new(
        id: "PARCH019",
        title: "Render attribute on a static member has no effect",
        messageFormat: "Member '{0}' is static, so '[{1}]' has no effect; the per-template maps that dispatch it walk instance members only. The value still binds and renders as plain text. Make the member non-static, or drop the attribute.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch019--render-attribute-on-a-static-member-has-no-effect");

    public static readonly DiagnosticDescriptor AmbiguousTemplate = new(
        id: "PARCH020",
        title: "Model matches more than one template file",
        messageFormat: "Model '{0}' matches more than one template file: {1}. A model binds exactly one '{0}.parchment.docx' or '{0}.parchment.md' — remove or rename the others.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch020--model-matches-more-than-one-template-file");

    public static readonly DiagnosticDescriptor EditableCollectionInvalid = new(
        id: "PARCH022",
        title: "[EditableField] collection has an unsupported shape",
        messageFormat: "Model '{0}' member '{1}' is an [EditableField] collection but {2}",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch022--editablefield-collection-has-an-unsupported-shape");

    public static readonly DiagnosticDescriptor ConflictingFormatMarkers = new(
        id: "PARCH023",
        title: "Member carries conflicting format markers",
        messageFormat: "Model '{0}' member '{1}' carries conflicting format markers — [Html], [Markdown] and [StringSyntax] disagree. Pick one; there is no principled winner to apply silently.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch023--member-carries-conflicting-format-markers");

    // An error, like every other "this format marker cannot be honoured" case — PARCH010, PARCH023,
    // PARCH007/008. The document still renders, but it renders the member's markup as visible text,
    // which is wrong output rather than a lesser rendering. A build that trips this has been
    // producing wrong documents already; stopping it is the point.
    public static readonly DiagnosticDescriptor FormatMarkerThroughLoopVariable = new(
        id: "PARCH024",
        title: "[Html]/[Markdown] member is reached through a loop variable",
        messageFormat: "Template '{0}' token '{1}' reaches the [Html]/[Markdown] member '{2}' through a loop variable. The format is resolved by walking the model from its root, which cannot address one iteration's item, so the value would render as text with its markup visible. Return a MarkdownToken or HtmlToken from the member instead, or apply the '| markdown' / '| html' filter at the token.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch024--html--markdown-member-reached-through-a-loop-variable");

    // The two say the same thing by different means, and applying both applies the conversion twice:
    // the token becomes a marker, and the marker's own text is then converted, so the marker renders
    // into the document. Refused rather than picking a winner, for the same reason as PARCH023.
    public static readonly DiagnosticDescriptor FormatMarkerOnTokenValue = new(
        id: "PARCH025",
        title: "Format marker on a member already typed as a token",
        messageFormat: "Model '{0}' member '{1}' is typed as a TokenValue and also carries [Html]/[Markdown]/[StringSyntax]. The type already selects the rendering, so the marker would apply it a second time — to the placeholder the first one produced, which renders that placeholder into the document. Drop one: keep the type, or keep the marker and type the member as a plain string.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch025--format-marker-on-a-member-already-typed-as-a-token");

    // The generated code calls the method directly, so an unresolvable name would otherwise
    // surface as a compile error inside a .g.cs file - correct, but pointing at code nobody wrote.
    // This reports it on the member that named it instead, and suppresses the emission so the
    // diagnostic is the only error rather than the first of a cascade.
    public static readonly DiagnosticDescriptor ExcelsiorConfigureMissing = new(
        id: "PARCH026",
        title: "[ExcelsiorTable] Configure method not found",
        messageFormat: "Model '{0}' member '{1}' names '{2}' in [ExcelsiorTable(Configure)], which is not a static method on '{3}'. Declare 'static void {2}(WordTableBuilder<TElement> builder)' beside the member, and reference it with nameof.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch026--excelsiortable-configure-method-not-found");

    public static readonly DiagnosticDescriptor AmbiguousStyleDoc = new(
        id: "PARCH021",
        title: "Model matches more than one style document",
        messageFormat: "Model '{0}' matches more than one style document: {1}. A markdown template takes at most one '{0}.parchment.dotx' — remove or rename the others.",
        category: "Parchment",
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true,
        helpLinkUri: "https://github.com/Papyrine/Parchment#parch021--model-matches-more-than-one-style-document");
}
