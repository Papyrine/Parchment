namespace Parchment.Generated;

/// <summary>
/// Public entry points called from source-generator-emitted <c>RegisterWith</c> helpers.
/// Pre-populates the runtime's per-type registration caches so the reflection-based
/// <see cref="SharedFluid.RegisterModel"/> / <c>*Map.Build</c> walks short-circuit when
/// <see cref="TemplateStore.RegisterDocxTemplate{TModel}(string, string, ProtectionMode)"/> runs.
///
/// Not intended for hand-written consumption — call sites are emitted by the
/// <c>Parchment.ParchmentModelAttribute</c> source generator. The runtime
/// <see cref="TemplateStore.RegisterDocxTemplate{TModel}(string, string, ProtectionMode)"/> path stays
/// fully functional for callers that can't use the source generator (POCO models, dynamic
/// template paths, etc.).
/// </summary>
public static class GeneratedRegistration
{
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
}