/// <summary>
/// Cache of string-typed properties reachable from a model type that are annotated to render as
/// html or markdown (via a user-defined <c>[Html]</c> / <c>[Markdown]</c> attribute or via
/// <c>[StringSyntax("html")]</c> / <c>[StringSyntax("markdown")]</c>). Built once at template
/// registration time; render-time lookup is a single dictionary hit.
/// </summary>
sealed class FormatMap
{
    readonly Dictionary<string, FormatKind> entries;

    FormatMap(Dictionary<string, FormatKind> entries) =>
        this.entries = entries;

    public bool IsEmpty => entries.Count == 0;

    public bool TryGet(string dottedPath, out FormatKind kind) =>
        entries.TryGetValue(dottedPath, out kind);

    public static FormatMap Build(Type modelType, string templateName)
    {
        var entries = new Dictionary<string, FormatKind>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<Type> { modelType };
        WalkType(modelType, [], entries, visited, templateName);
        return new(entries);
    }

    static void WalkType(
        Type type,
        List<string> pathSegments,
        Dictionary<string, FormatKind> entries,
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
                entries[dottedPath] = format.Value;
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

            WalkType(propUnderlying, nextSegments, entries, visited, templateName);
            visited.Remove(propUnderlying);
        }
    }

    static FormatKind? DetectFormat(PropertyInfo property, string templateName, Type owner)
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
            return FormatKind.Html;
        }

        if (hasMarkdownAttribute || syntax == "markdown")
        {
            return FormatKind.Markdown;
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
            type == typeof(TimeSpan) ||
            type == typeof(Guid) ||
            type == typeof(Uri))
        {
            return false;
        }

        return !typeof(IEnumerable).IsAssignableFrom(type);
    }
}

enum FormatKind
{
    Html,
    Markdown
}
