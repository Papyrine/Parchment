/// <summary>
/// Cache of <see cref="IEnumerable{String}"/>-typed properties reachable from a model type,
/// keyed by their dotted path from the root (e.g. <c>Customer.Tags</c>). Built once at template
/// registration time so render-time lookup is a single dictionary hit. Detection is type-driven
/// (no attribute) — any property assignable to <c>IEnumerable&lt;string&gt;</c> qualifies, except
/// properties already marked <c>[ExcelsiorTable]</c> (those keep ownership via the Excelsior path).
/// </summary>
sealed class StringListMap
{
    static ConcurrentDictionary<Type, StringListMap> precompiledCache = new();

    Dictionary<string, Func<object, object?>> entries;

    StringListMap(Dictionary<string, Func<object, object?>> entries) =>
        this.entries = entries;

    public static StringListMap Empty { get; } = new(new(StringComparer.OrdinalIgnoreCase));

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out Func<object, object?>? getter) =>
        entries.TryGetValue(dottedPath, out getter);

    // No reflection fallback — see ExcelsiorTableMap.Build.
    public static StringListMap Build(Type modelType) =>
        precompiledCache.GetValueOrDefault(modelType, Empty);

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<StringListMapEntry> entries)
    {
        var dict = new Dictionary<string, Func<object, object?>>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = entry.Getter;
        }

        precompiledCache[modelType] = new(dict);
    }
}
