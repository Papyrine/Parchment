namespace Parchment.Generated;

/// <summary>
/// Template definitions registered by the source generator's module initializers. A definition is
/// the raw template content (docx bytes, markdown text, optional style-source bytes) keyed by its
/// model type. Registration into a <see cref="TemplateStore"/> is deferred until first render so
/// module load stays cheap and per-store settings (image policies, page numbers) apply.
/// </summary>
static class GeneratedTemplateDefinitions
{
    static readonly ConcurrentDictionary<Type, TemplateDefinition> definitions = new();

    public static void Add(Type modelType, TemplateDefinition definition) =>
        definitions[modelType] = definition;

    public static bool TryGet(Type modelType, [NotNullWhen(true)] out TemplateDefinition? definition) =>
        definitions.TryGetValue(modelType, out definition);
}

abstract class TemplateDefinition;

sealed class DocxTemplateDefinition(byte[] template, ProtectionMode protection) :
    TemplateDefinition
{
    public byte[] Template { get; } = template;
    public ProtectionMode Protection { get; } = protection;
}

sealed class MarkdownTemplateDefinition(string markdown, byte[]? styleSource) :
    TemplateDefinition
{
    public string Markdown { get; } = markdown;
    public byte[]? StyleSource { get; } = styleSource;
}
