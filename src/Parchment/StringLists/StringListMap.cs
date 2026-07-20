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

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out Func<object, object?>? getter) =>
        entries.TryGetValue(dottedPath, out getter);

    public static StringListMap Build(Type modelType)
    {
        if (precompiledCache.TryGetValue(modelType, out var cached))
        {
            return cached;
        }

        var entries = new Dictionary<string, Func<object, object?>>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<Type> { modelType };
        WalkType(modelType, [], static root => root, entries, visited);
        return new(entries);
    }

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<StringListMapEntry> entries)
    {
        var dict = new Dictionary<string, Func<object, object?>>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = entry.Getter;
        }

        precompiledCache[modelType] = new(dict);
    }

    static void WalkType(
        Type type,
        List<string> pathSegments,
        Func<object, object?> getter,
        Dictionary<string, Func<object, object?>> entries,
        HashSet<Type> visited)
    {
        foreach (var (name, memberType, memberGetter, excelsior) in ExcelsiorTableMap.EnumerateMembers(type))
        {
            // [ExcelsiorTable] keeps full ownership of the member — don't shadow it with the
            // string-list path even if the element type happens to be string.
            if (excelsior != null)
            {
                continue;
            }

            var nextSegments = new List<string>(pathSegments) { name };
            var nextGetter = ChainGetter(getter, memberGetter);
            var underlying = Nullable.GetUnderlyingType(memberType) ?? memberType;

            if (IsEnumerableOfString(underlying))
            {
                var dottedPath = string.Join('.', nextSegments);
                entries[dottedPath] = nextGetter;
                // Don't descend into a string-list leaf.
                continue;
            }

            // Descend into POCO members. Skip leaves and other collection types (loops handle
            // those). Track visited types per branch so self-referential models don't recurse
            // forever; the same type can still appear at multiple unrelated paths.
            if (!ModelGraph.ShouldDescend(underlying))
            {
                continue;
            }

            if (!visited.Add(underlying))
            {
                continue;
            }

            WalkType(underlying, nextSegments, nextGetter, entries, visited);
            visited.Remove(underlying);
        }
    }

    static bool IsEnumerableOfString(Type type) =>
        typeof(IEnumerable<string>).IsAssignableFrom(type);

    static Func<object, object?> ChainGetter(Func<object, object?> upstream, Func<object, object?> memberGetter) =>
        root =>
        {
            var parent = upstream(root);
            if (parent == null)
            {
                return null;
            }

            return memberGetter(parent);
        };

}
