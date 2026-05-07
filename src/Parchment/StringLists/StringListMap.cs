/// <summary>
/// Cache of <see cref="IEnumerable{String}"/>-typed properties reachable from a model type,
/// keyed by their dotted path from the root (e.g. <c>Customer.Tags</c>). Built once at template
/// registration time so render-time lookup is a single dictionary hit. Detection is type-driven
/// (no attribute) — any property assignable to <c>IEnumerable&lt;string&gt;</c> qualifies, except
/// properties already marked <c>[ExcelsiorTable]</c> (those keep ownership via the Excelsior path).
/// </summary>
sealed class StringListMap
{
    readonly Dictionary<string, Func<object, object?>> entries;

    StringListMap(Dictionary<string, Func<object, object?>> entries) =>
        this.entries = entries;

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out Func<object, object?>? getter) =>
        entries.TryGetValue(dottedPath, out getter);

    public static StringListMap Build(Type modelType)
    {
        var entries = new Dictionary<string, Func<object, object?>>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<Type> { modelType };
        WalkType(modelType, [], static root => root, entries, visited);
        return new(entries);
    }

    static void WalkType(
        Type type,
        List<string> pathSegments,
        Func<object, object?> getter,
        Dictionary<string, Func<object, object?>> entries,
        HashSet<Type> visited)
    {
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanRead)
            {
                continue;
            }

            // [ExcelsiorTable] keeps full ownership of the property — don't shadow it with the
            // string-list path even if the element type happens to be string.
            if (property.GetCustomAttribute<ExcelsiorTableAttribute>() != null)
            {
                continue;
            }

            var nextSegments = new List<string>(pathSegments) { property.Name };
            var nextGetter = ChainGetter(getter, property);
            var underlying = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;

            if (IsEnumerableOfString(underlying))
            {
                var dottedPath = string.Join('.', nextSegments);
                entries[dottedPath] = nextGetter;
                // Don't descend into a string-list leaf.
                continue;
            }

            // Descend into POCO properties. Skip leaves and other collection types (loops handle
            // those). Track visited types per branch so self-referential models don't recurse
            // forever; the same type can still appear at multiple unrelated paths.
            if (!ShouldDescend(underlying))
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

    static Func<object, object?> ChainGetter(Func<object, object?> upstream, PropertyInfo property) =>
        root =>
        {
            var parent = upstream(root);
            if (parent == null)
            {
                return null;
            }

            return property.GetValue(parent);
        };

    static bool ShouldDescend(Type type)
    {
        if (type.IsPrimitive || type.IsEnum)
        {
            return false;
        }

        if (type == typeof(string) ||
            type == typeof(decimal) ||
            type == typeof(DateTime) ||
            type == typeof(DateTimeOffset) ||
            type == typeof(TimeSpan) ||
            type == typeof(Guid) ||
            type == typeof(Uri))
        {
            return false;
        }

        return !typeof(IEnumerable).IsAssignableFrom(type);
    }
}
