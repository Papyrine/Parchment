/// <summary>
/// Cache of string-typed properties reachable from a model type that are annotated to render as
/// html or markdown (via a user-defined <c>[Html]</c> / <c>[Markdown]</c> attribute or via
/// <c>[StringSyntax("html")]</c> / <c>[StringSyntax("markdown")]</c>). Built once at template
/// registration time; render-time lookup is a single dictionary hit.
/// </summary>
sealed class FormatMap
{
    static ConcurrentDictionary<Type, FormatMap> precompiledCache = new();

    Dictionary<string, FormatEntry> entries;

    FormatMap(Dictionary<string, FormatEntry> entries) =>
        this.entries = entries;

    public static FormatMap Empty { get; } = new(new(StringComparer.OrdinalIgnoreCase));

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out FormatEntry? entry) =>
        entries.TryGetValue(dottedPath, out entry);

    // No reflection fallback — see ExcelsiorTableMap.Build.
    public static FormatMap Build(Type modelType) =>
        precompiledCache.GetValueOrDefault(modelType, Empty);

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<FormatMapEntry> entries)
    {
        var dict = new Dictionary<string, FormatEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = new(entry.Format, entry.Getter);
        }

        precompiledCache[modelType] = new(dict);
    }

}
