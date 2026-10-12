/// <summary>
/// Cache of <see cref="ExcelsiorTableAttribute"/>-marked properties reachable from a model type,
/// keyed by their dotted path from the root (e.g. <c>Customer.Lines</c>). Built once at template
/// registration time so render-time lookup is a single dictionary hit.
/// </summary>
sealed class ExcelsiorTableMap
{
    static ConcurrentDictionary<Type, ExcelsiorTableMap> precompiledCache = new();

    Dictionary<string, ExcelsiorTableEntry> entries;

    ExcelsiorTableMap(Dictionary<string, ExcelsiorTableEntry> entries) =>
        this.entries = entries;

    public static ExcelsiorTableMap Empty { get; } = new(new(StringComparer.OrdinalIgnoreCase));

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out ExcelsiorTableEntry? entry) =>
        entries.TryGetValue(dottedPath, out entry);

    // No reflection fallback: the source generator walks the model graph at compile time and
    // registers the entries via GeneratedRegistration. A model with no [ExcelsiorTable] members
    // simply has no registration, which is the empty map.
    public static ExcelsiorTableMap Build(Type modelType) =>
        precompiledCache.GetValueOrDefault(modelType, Empty);

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<ExcelsiorTableMapEntry> entries)
    {
        var dict = new Dictionary<string, ExcelsiorTableEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = new(entry.ElementType, entry.Getter, entry.HeadingParagraphStyle, entry.BodyParagraphStyle, entry.TableStyle, entry.Configure);
        }

        precompiledCache[modelType] = new(dict);
    }
}