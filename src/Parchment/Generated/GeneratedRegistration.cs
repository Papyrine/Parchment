namespace Parchment.Generated;

/// <summary>
/// Public entry points called from source-generator-emitted module initializers. Stores each
/// model's embedded template definition, and populates the runtime's per-type registration
/// caches — the pre-compiled accessors and per-template maps every registration requires. There
/// is no reflection fallback: a model registered by hand must carry <c>[ParchmentBindable]</c>
/// (or <c>[ParchmentModel]</c>) so these calls have run when its assembly loaded.
///
/// Not intended for hand-written consumption — call sites are emitted by the Parchment source
/// generator.
/// </summary>
public static class GeneratedRegistration
{
    /// <summary>
    /// Stores the embedded docx template for <paramref name="modelType"/>. Called from the
    /// generated module initializer; every <see cref="TemplateStore"/> materializes the definition
    /// on the model's first render.
    /// </summary>
    public static void RegisterDocxTemplate(Type modelType, byte[] template, ProtectionMode protection = ProtectionMode.WhenEditable)
    {
        ArgumentNullException.ThrowIfNull(template);
        GeneratedTemplateDefinitions.Add(modelType, new DocxTemplateDefinition(() => template, protection));
    }

    /// <summary>
    /// Stores the embedded docx template for <paramref name="modelType"/>, whose bytes are the
    /// <see cref="EmbeddedBytesAttribute"/> of <paramref name="embedded"/>. The attribute is read
    /// when a <see cref="TemplateStore"/> materializes the definition, not here: a model that is
    /// never rendered costs nothing beyond this call.
    /// </summary>
    public static void RegisterDocxTemplate(Type modelType, Type embedded, ProtectionMode protection = ProtectionMode.WhenEditable)
    {
        ArgumentNullException.ThrowIfNull(embedded);
        GeneratedTemplateDefinitions.Add(modelType, new DocxTemplateDefinition(() => ReadEmbeddedBytes(embedded), protection));
    }

    /// <summary>
    /// Stores the embedded markdown template — and its resolved style source, when the project
    /// carries one — for <paramref name="modelType"/>. Called from the generated module
    /// initializer; every <see cref="TemplateStore"/> materializes the definition on the model's
    /// first render.
    /// </summary>
    public static void RegisterMarkdownTemplate(Type modelType, string markdown, byte[]? styleSource = null)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        if (styleSource == null)
        {
            GeneratedTemplateDefinitions.Add(modelType, new MarkdownTemplateDefinition(markdown, null));
            return;
        }

        GeneratedTemplateDefinitions.Add(modelType, new MarkdownTemplateDefinition(markdown, () => styleSource));
    }

    /// <summary>
    /// Stores the embedded markdown template for <paramref name="modelType"/> with a style source
    /// whose bytes are the <see cref="EmbeddedBytesAttribute"/> of <paramref name="embeddedStyle"/>,
    /// read when a <see cref="TemplateStore"/> materializes the definition.
    /// </summary>
    public static void RegisterMarkdownTemplate(Type modelType, string markdown, Type embeddedStyle)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        ArgumentNullException.ThrowIfNull(embeddedStyle);
        GeneratedTemplateDefinitions.Add(modelType, new MarkdownTemplateDefinition(markdown, () => ReadEmbeddedBytes(embeddedStyle)));
    }

    static byte[] ReadEmbeddedBytes(Type embedded)
    {
        var attribute = embedded.GetCustomAttribute<EmbeddedBytesAttribute>(false);
        if (attribute == null)
        {
            throw new InvalidOperationException($"{embedded.FullName} carries no [EmbeddedBytes] attribute.");
        }

        return Convert.FromBase64String(attribute.Base64);
    }

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