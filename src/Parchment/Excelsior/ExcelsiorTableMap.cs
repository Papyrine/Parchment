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

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out ExcelsiorTableEntry? entry) =>
        entries.TryGetValue(dottedPath, out entry);

    public static ExcelsiorTableMap Build(Type modelType, string templateName)
    {
        if (precompiledCache.TryGetValue(modelType, out var cached))
        {
            return cached;
        }

        var entries = new Dictionary<string, ExcelsiorTableEntry>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<Type>
        {
            modelType
        };
        WalkType(modelType, [], static root => root, entries, visited, templateName);
        return new(entries);
    }

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<ExcelsiorTableMapEntry> entries)
    {
        var dict = new Dictionary<string, ExcelsiorTableEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = new(entry.ElementType, entry.Getter, entry.HeadingParagraphStyle, entry.BodyParagraphStyle);
        }

        precompiledCache[modelType] = new(dict);
    }

    static void WalkType(
        Type type,
        List<string> pathSegments,
        Func<object, object?> getter,
        Dictionary<string, ExcelsiorTableEntry> entries,
        HashSet<Type> visited,
        string templateName)
    {
        foreach (var (name, memberType, memberGetter, excelsior) in EnumerateMembers(type))
        {
            var nextSegments = new List<string>(pathSegments)
            {
                name
            };
            var nextGetter = ChainGetter(getter, memberGetter);

            if (excelsior != null)
            {
                var elementType = ModelValidator.TryResolveElementType(memberType);
                if (elementType == null ||
                    elementType == typeof(char))
                {
                    throw new ParchmentRegistrationException(
                        templateName,
                        $"[ExcelsiorTable] member '{type.Name}.{name}' must be an IEnumerable<T> (excluding string).");
                }

                var dottedPath = string.Join('.', nextSegments);
                entries[dottedPath] = new(elementType, nextGetter, excelsior.HeadingParagraphStyle, excelsior.BodyParagraphStyle);
                // Don't descend into the element type — collection items become Excelsior columns.
                continue;
            }

            // Descend into POCO members looking for nested [ExcelsiorTable] members. Skip leaves
            // (primitives, strings, dates, etc.) and collection types (loops handle those). Track
            // visited types per branch so self-referential models don't recurse forever; the same
            // type can still appear at multiple unrelated paths.
            var underlying = Nullable.GetUnderlyingType(memberType) ?? memberType;
            if (!ModelGraph.ShouldDescend(underlying))
            {
                continue;
            }

            if (!visited.Add(underlying))
            {
                continue;
            }

            WalkType(underlying, nextSegments, nextGetter, entries, visited, templateName);
            visited.Remove(underlying);
        }
    }

    internal static IEnumerable<(string Name, Type Type, Func<object, object?> Getter, ExcelsiorTableAttribute? ExcelsiorTable)> EnumerateMembers(Type type)
    {
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanRead)
            {
                continue;
            }

            yield return (
                property.Name,
                property.PropertyType,
                property.GetValue,
                property.GetCustomAttribute<ExcelsiorTableAttribute>());
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            yield return (
                field.Name,
                field.FieldType,
                field.GetValue,
                field.GetCustomAttribute<ExcelsiorTableAttribute>());
        }
    }

    static Func<object, object?> ChainGetter(Func<object, object?> upstream, Func<object, object?> memberGetter) =>
        root =>
        {
            var parent = upstream(root);
            return parent == null ? null : memberGetter(parent);
        };

}