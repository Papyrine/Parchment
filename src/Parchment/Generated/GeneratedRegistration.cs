namespace Parchment.Generated;

/// <summary>
/// Public entry points called from source-generator-emitted module initializers. Stores each
/// model's embedded template definition, and pre-populates the runtime's per-type registration
/// caches so the reflection-based <see cref="SharedFluid.RegisterModel"/> / <c>*Map.Build</c>
/// walks short-circuit when a <see cref="TemplateStore"/> materializes the template.
///
/// Not intended for hand-written consumption — call sites are emitted by the
/// <c>Parchment.ParchmentModelAttribute</c> source generator. The runtime
/// <see cref="TemplateStore.RegisterDocxTemplate{TModel}(string, ProtectionMode)"/> path stays
/// fully functional for callers that can't use the source generator (POCO models, dynamic
/// template paths, etc.).
/// </summary>
public static class GeneratedRegistration
{
    /// <summary>
    /// Stores the embedded docx template for <paramref name="modelType"/>. Called from the
    /// generated module initializer; every <see cref="TemplateStore"/> materializes the definition
    /// on the model's first render.
    /// </summary>
    public static void RegisterDocxTemplate(Type modelType, byte[] template, ProtectionMode protection = ProtectionMode.WhenEditable) =>
        GeneratedTemplateDefinitions.Add(modelType, new DocxTemplateDefinition(template, protection));

    /// <summary>
    /// Stores the embedded markdown template — and its resolved style source, when the project
    /// carries one — for <paramref name="modelType"/>. Called from the generated module
    /// initializer; every <see cref="TemplateStore"/> materializes the definition on the model's
    /// first render.
    /// </summary>
    public static void RegisterMarkdownTemplate(Type modelType, string markdown, byte[]? styleSource = null) =>
        GeneratedTemplateDefinitions.Add(modelType, new MarkdownTemplateDefinition(markdown, styleSource));

    public static void RegisterFluidAccessors(
        Type type,
        IEnumerable<KeyValuePair<string, IMemberAccessor>> accessors) =>
        SharedFluid.RegisterPrecompiledAccessors(type, accessors);

    public static void RegisterExcelsiorTable(
        Type modelType,
        IEnumerable<ExcelsiorTableMapEntry> entries) =>
        ExcelsiorTableMap.RegisterPrecompiled(modelType, entries);

    public static void RegisterFormat(
        Type modelType,
        IEnumerable<FormatMapEntry> entries) =>
        FormatMap.RegisterPrecompiled(modelType, entries);

    public static void RegisterStringList(
        Type modelType,
        IEnumerable<StringListMapEntry> entries) =>
        StringListMap.RegisterPrecompiled(modelType, entries);

    public static void RegisterEditable(
        Type modelType,
        IEnumerable<EditableFieldMapEntry> entries) =>
        EditableMap.RegisterPrecompiled(modelType, entries);

    public static void RegisterEditableCollections(
        Type modelType,
        IEnumerable<CollectionFieldMapEntry> entries) =>
        EditableMap.RegisterPrecompiledCollections(modelType, entries);
}