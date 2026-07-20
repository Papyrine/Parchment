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
        var visited = new HashSet<Type>
        {
            modelType
        };
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
        foreach (var (name, memberType, memberGetter, member) in EnumerateMembers(type))
        {
            var nextSegments = new List<string>(pathSegments)
            {
                name
            };
            var nextGetter = ChainGetter(getter, memberGetter);

            var format = DetectFormat(member, templateName, type);
            if (format != null)
            {
                // A member that is also [EditableField] is claimed by the editable map — it renders
                // as an editable rich-content control and extracts back to HTML, not the read-only
                // format path. (EditableMap mirrors this: it rejects [Markdown]+[EditableField] and
                // accepts [Html]+[EditableField].)
                if (member.GetCustomAttribute<EditableFieldAttribute>(true) != null)
                {
                    continue;
                }

                var underlying = Nullable.GetUnderlyingType(memberType) ?? memberType;
                if (underlying == typeof(string))
                {
                    var dottedPath = string.Join('.', nextSegments);
                    entries[dottedPath] = new(format.Value, nextGetter);
                    continue;
                }

                throw new ParchmentRegistrationException(
                    templateName,
                    $"[{format}] member '{type.Name}.{name}' must be a string.");
            }

            var memberUnderlying = Nullable.GetUnderlyingType(memberType) ?? memberType;
            if (!ModelGraph.ShouldDescend(memberUnderlying))
            {
                continue;
            }

            if (!visited.Add(memberUnderlying))
            {
                continue;
            }

            WalkType(memberUnderlying, nextSegments, nextGetter, entries, visited, templateName);
            visited.Remove(memberUnderlying);
        }
    }

    internal static IEnumerable<(string Name, Type Type, Func<object, object?> Getter, MemberInfo Member)> EnumerateMembers(Type type)
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
                property);
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            yield return (
                field.Name,
                field.FieldType,
                field.GetValue,
                field);
        }
    }

    static Func<object, object?> ChainGetter(Func<object, object?> upstream, Func<object, object?> memberGetter) =>
        root =>
        {
            var parent = upstream(root);
            return parent == null ? null : memberGetter(parent);
        };

    static FormatMapKind? DetectFormat(MemberInfo member, string templateName, Type owner)
    {
        var hasHtmlAttribute = false;
        var hasMarkdownAttribute = false;
        foreach (var attribute in member.GetCustomAttributes(true))
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

        var syntax = ReadStringSyntax(member);

        if (hasHtmlAttribute && hasMarkdownAttribute)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Member '{owner.Name}.{member.Name}': cannot have both [Html] and [Markdown].");
        }

        if (hasHtmlAttribute && syntax == "markdown")
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Member '{owner.Name}.{member.Name}': mismatched format — [Html] contradicts [StringSyntax(\"markdown\")].");
        }

        if (hasMarkdownAttribute &&
            syntax == "html")
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"Member '{owner.Name}.{member.Name}': mismatched format — [Markdown] contradicts [StringSyntax(\"html\")].");
        }

        if (hasHtmlAttribute ||
            syntax == "html")
        {
            return FormatMapKind.Html;
        }

        if (hasMarkdownAttribute ||
            syntax == "markdown")
        {
            return FormatMapKind.Markdown;
        }

        return null;
    }

    static string? ReadStringSyntax(MemberInfo member) =>
        member.GetCustomAttribute<StringSyntaxAttribute>(true)?.Syntax.ToLowerInvariant();

}

