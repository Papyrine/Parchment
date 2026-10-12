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
        if (!TryResolveRoot(modelType, path.Root, scope, out var currentType))
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
            if (!TryResolveSegment(currentType, segment, out var next))
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    MissingMemberMessage(currentType, segment),
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

    public static bool TryResolveElementType(Type enumerableType, [NotNullWhen(true)] out Type? elementType)
    {
        if (enumerableType.IsArray)
        {
            elementType = enumerableType.GetElementType();
            return elementType != null;
        }

        foreach (var i in enumerableType.GetInterfaces())
        {
            if (i.IsGenericType &&
                i.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            {
                elementType = i.GetGenericArguments()[0];
                return true;
            }
        }

        if (enumerableType.IsGenericType &&
            enumerableType.GetGenericTypeDefinition() == typeof(IEnumerable<>))
        {
            elementType = enumerableType.GetGenericArguments()[0];
            return true;
        }

        elementType = null;
        return false;
    }

    static string MissingMemberMessage(Type type, string name)
    {
        var kind = LiquidMembers.Classify(type);
        if (kind != LiquidKind.None)
        {
            return LiquidMembers.MissingMessage(kind, type, name);
        }

        if (HasNoAccessors(type))
        {
            return $"Member '{name}' of '{type.Name}' is not available to a template. " +
                   "No member of a system type is bound; expose the value as a member of the model instead.";
        }

        return $"Member '{name}' is not a property or field of '{type.Name}'";
    }

    /// <summary>
    /// Resolves a segment after the first. Unlike the root, which binds through the model's own
    /// accessors, what such a segment is read from may be a string or a collection, and those
    /// answer to liquid's members rather than their CLR ones — see <see cref="LiquidMembers"/>.
    /// It may also be a system type, which has no accessors at all.
    /// </summary>
    public static bool TryResolveSegment(Type type, string name, [NotNullWhen(true)] out Type? memberType)
    {
        var kind = LiquidMembers.Classify(type);
        if (kind != LiquidKind.None)
        {
            return LiquidMembers.TryResolve(kind, type, name, out memberType);
        }

        if (HasNoAccessors(type))
        {
            memberType = null;
            return false;
        }

        return TryResolveMember(type, name, out memberType);
    }

    // The generator emits accessors for the model's own types and stops at a system one
    // (ShapeBuilder.IsSystemType): Year on a DateTime, Value on a Nullable<T> and Item1 on a tuple
    // are all there to reflect on, and all render blank. KeyValuePair<K, V> is the one system type
    // it surfaces members for, Key and Value, which are also all that type declares.
    static bool HasNoAccessors(Type type)
    {
        if (type.IsGenericType &&
            type.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
        {
            return false;
        }

        var ns = type.Namespace;
        if (ns == null)
        {
            return false;
        }

        return ns == "System" ||
               ns.StartsWith("System.", StringComparison.Ordinal);
    }

    public static bool TryResolveMember(Type type, string name, [NotNullWhen(true)] out Type? memberType)
    {
        var property = type.GetProperty(
            name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.IgnoreCase);
        // Readable in public, and not an indexer: the generator emits an accessor for nothing else
        // (ShapeBuilder.TryGetMemberType), so anything else would register and then render blank.
        if (property is { GetMethod.IsPublic: true } &&
            property.GetIndexParameters().Length == 0)
        {
            memberType = property.PropertyType;
            return true;
        }

        var field = type.GetField(
            name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.IgnoreCase);
        memberType = field?.FieldType;
        return memberType != null;
    }

    static bool TryResolveRoot(
        Type modelType,
        string name,
        IReadOnlyDictionary<string, Type>? scope,
        [NotNullWhen(true)] out Type? rootType)
    {
        if (scope != null &&
            scope.TryGetValue(name, out var scoped))
        {
            rootType = scoped;
            return true;
        }

        // Nothing else names the model: it is the scope a root identifier is read from, so a
        // `model.` or type-name prefix is one more unknown member, and renders blank.
        return TryResolveMember(modelType, name, out rootType);
    }
}
