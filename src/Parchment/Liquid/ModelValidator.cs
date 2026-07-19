/// <summary>
/// Reflection-based validator that walks an <see cref="IdentifierPath"/> against a .NET type
/// to ensure every segment resolves to a real member. Used by the runtime registration path.
/// </summary>
static class ModelValidator
{
    public static void Validate(
        Type modelType,
        IdentifierPath path,
        IReadOnlyDictionary<string, Type>? scope,
        string templateName,
        string? partUri,
        string? tokenSource)
    {
        var currentType = ResolveRoot(modelType, path.Root, scope);
        if (currentType == null)
        {
            throw new ParchmentRegistrationException(
                templateName,
                UnresolvedRootMessage(path.Root, modelType),
                partUri,
                tokenSource,
                path.ToString());
        }

        for (var i = 1; i < path.Segments.Count; i++)
        {
            var segment = path.Segments[i];
            var next = ResolveMember(currentType, segment);
            if (next == null)
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"Member '{segment}' is not a property or field of '{currentType.Name}'",
                    partUri,
                    tokenSource,
                    path.ToString());
            }

            currentType = next;
        }
    }

    // `forloop` resolves against no model member by design, so the generic "not a member of" text
    // reads as a missing property and sends the reader looking for one. The markdown flow allows it
    // inside a loop body; reaching here means either a docx template, which runs its own loop scope
    // and never populates it, or a markdown template referencing it outside any loop.
    static string UnresolvedRootMessage(string root, Type modelType)
    {
        if (root == "forloop")
        {
            return "'forloop' is provided by liquid only inside a {% for %} body of a markdown " +
                   "template. Docx templates iterate through their own scope tree and do not " +
                   "provide it.";
        }

        return $"Identifier '{root}' is not a member of '{modelType.Name}'";
    }

    public static Type? TryResolveElementType(Type enumerableType)
    {
        if (enumerableType.IsArray)
        {
            return enumerableType.GetElementType();
        }

        foreach (var i in enumerableType.GetInterfaces())
        {
            if (i.IsGenericType &&
                i.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            {
                return i.GetGenericArguments()[0];
            }
        }

        if (enumerableType.IsGenericType &&
            enumerableType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
        {
            return enumerableType.GetGenericArguments()[0];
        }

        return null;
    }

    public static Type? ResolveMember(Type type, string name)
    {
        var property = type.GetProperty(
            name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.IgnoreCase);
        if (property != null)
        {
            return property.PropertyType;
        }

        var field = type.GetField(
            name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.IgnoreCase);
        return field?.FieldType;
    }

    static Type? ResolveRoot(Type modelType, string name, IReadOnlyDictionary<string, Type>? scope)
    {
        if (scope != null &&
            scope.TryGetValue(name, out var scoped))
        {
            return scoped;
        }

        var type = ResolveMember(modelType, name);
        if (type != null)
        {
            return type;
        }

        if (MatchesRootIdentifier(modelType, name))
        {
            return modelType;
        }

        return null;
    }

    static bool MatchesRootIdentifier(Type modelType, string name) =>
        string.Equals(name, modelType.Name, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "model", StringComparison.OrdinalIgnoreCase);
}
