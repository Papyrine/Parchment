/// <summary>
/// Cache of string-typed properties reachable from a model type that are annotated to render as
/// html or markdown (via a user-defined <c>[Html]</c> / <c>[Markdown]</c> attribute or via
/// <c>[StringSyntax("html")]</c> / <c>[StringSyntax("markdown")]</c>). Built once at template
/// registration time; render-time lookup is a single dictionary hit.
/// </summary>
sealed class FormatMap
{
    static readonly ConcurrentDictionary<Type, FormatMap> precompiledCache = new();

    readonly Dictionary<string, FormatEntry> entries;

    FormatMap(Dictionary<string, FormatEntry> entries) =>
        this.entries = entries;

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out FormatEntry? entry) =>
        entries.TryGetValue(dottedPath, out entry);

    public static FormatMap Build(Type modelType, string templateName)
    {
        if (precompiledCache.TryGetValue(modelType, out var cached))
        {
            return cached;
        }

        var entries = new Dictionary<string, FormatEntry>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<Type> { modelType };
        WalkType(modelType, [], static root => root, entries, visited, templateName);
        return new(entries);
    }

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<FormatMapEntry> entries)
    {
        var dict = new Dictionary<string, FormatEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = new(entry.Format, entry.Getter);
        }

        precompiledCache[modelType] = new(dict);
    }

    static void WalkType(
        Type type,
        List<string> pathSegments,
        Func<object, object?> getter,
        Dictionary<string, FormatEntry> entries,
        HashSet<Type> visited,
        string templateName)
    {
        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (!property.CanRead)
            {
                continue;
            }

            var nextSegments = new List<string>(pathSegments) { property.Name };
            var nextGetter = ChainGetter(getter, property);

            var format = DetectFormat(property, templateName, type);
            if (format != null)
            {
                var underlying = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
                if (underlying != typeof(string))
                {
                    throw new ParchmentRegistrationException(
                        templateName,
                        $"[{format}] property '{type.Name}.{property.Name}' must be a string.");
                }

                var dottedPath = string.Join('.', nextSegments);
                entries[dottedPath] = new(format.Value, nextGetter);
                continue;
            }

            var propUnderlying = Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType;
            if (!ShouldDescend(propUnderlying))
            {
                continue;
            }

            if (!visited.Add(propUnderlying))
            {
                continue;
            }

            WalkType(propUnderlying, nextSegments, nextGetter, entries, visited, templateName);
            visited.Remove(propUnderlying);
        }
    }

    static Func<object, object?> ChainGetter(Func<object, object?> upstream, PropertyInfo property) =>
        root =>
        {
            var parent = upstream(root);
            return parent == null ? null : property.GetValue(parent);
        };

    static FormatMapKind? DetectFormat(PropertyInfo property, string templateName, Type owner)
    {
        var hasHtmlAttribute = false;
        var hasMarkdownAttribute = false;
        foreach (var attribute in property.GetCustomAttributes(true))
        {
            var name = attribute.GetType().Name;
            if (name == "HtmlAttribute")
            {
                hasHtmlAttribute = true;
            }
            else if (name == "MarkdownAttribute")
            {
                hasMarkdownAttribute = true;
            }
        }

        var syntax = ReadStringSyntax(property);

        if (hasHtmlAttribute && hasMarkdownAttribute)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Property '{owner.Name}.{property.Name}': cannot have both [Html] and [Markdown].");
        }

        if (hasHtmlAttribute && syntax == "markdown")
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Property '{owner.Name}.{property.Name}': mismatched format — [Html] contradicts [StringSyntax(\"markdown\")].");
        }

        if (hasMarkdownAttribute && syntax == "html")
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Property '{owner.Name}.{property.Name}': mismatched format — [Markdown] contradicts [StringSyntax(\"html\")].");
        }

        if (hasHtmlAttribute || syntax == "html")
        {
            return FormatMapKind.Html;
        }

        if (hasMarkdownAttribute || syntax == "markdown")
        {
            return FormatMapKind.Markdown;
        }

        return null;
    }

    static string? ReadStringSyntax(PropertyInfo property)
    {
        foreach (var attribute in property.GetCustomAttributes(true))
        {
            if (attribute.GetType().FullName != "System.Diagnostics.CodeAnalysis.StringSyntaxAttribute")
            {
                continue;
            }

            var syntaxProperty = attribute.GetType().GetProperty("Syntax");
            if (syntaxProperty?.GetValue(attribute) is string value)
            {
                return value.ToLowerInvariant();
            }
        }

        return null;
    }

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
            type == typeof(Date) ||
            type == typeof(Time) ||
            type == typeof(TimeSpan) ||
            type == typeof(Guid) ||
            type == typeof(Uri))
        {
            return false;
        }

        return !typeof(IEnumerable).IsAssignableFrom(type);
    }
}

