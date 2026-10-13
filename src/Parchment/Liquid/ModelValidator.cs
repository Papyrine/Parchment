/// <summary>
/// Reflection-based validator that walks an <see cref="IdentifierPath"/> against a .NET type
/// to ensure every segment resolves to a real member. Used by the runtime registration path.
/// </summary>
[RequiresUnreferencedCode(Unsupported.Trimming)]
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
            if (segment == IdentifierPath.Index)
            {
                // Past an index the type cannot see through, nothing is known to check against.
                if (!LiquidMembers.TryResolveIndex(currentType, out var element))
                {
                    return;
                }

                currentType = element;
                continue;
            }

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

        // A model that is itself a collection is handed to the template as an object (see
        // SharedFluid.CreateContext), so neither what its collection base declares nor liquid's
        // array members exist on it. Both read as if they should, which the plain text below
        // would not explain.
        if (LiquidMembers.Classify(modelType) != LiquidKind.None &&
            (root is "size" or "first" or "last" ||
             TryResolveMember(modelType, root, out _)))
        {
            return $"Identifier '{root}' is not available on '{modelType.Name}'. '{modelType.Name}' is the model itself, " +
                   "which a template reads as an object, not as a collection: only the members the model " +
                   "declares bind, not those of its collection base type nor liquid's 'size', 'first' and " +
                   "'last'. Expose the value, or the items, through a member of the model.";
        }

        return $"Identifier '{root}' is not a member of '{modelType.Name}'";
    }

    public static bool TryResolveElementType(Type enumerableType, [NotNullWhen(true)] out Type? elementType)
    {
        // A string is an IEnumerable<char>, but liquid does not walk its characters: a loop over one
        // runs once with the whole string. Nothing iterates it as a collection, so it is not one
        // here. Lockstep with the source generator's ShapeBuilder.BuildEntry, which reports PARCH002.
        if (enumerableType == typeof(string))
        {
            elementType = null;
            return false;
        }

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
        if (IsTokenValue(type))
        {
            return $"Member '{name}' of '{type.Name}' is not available to a template. " +
                   "A token value is rendered whole and none of its members is bound; expose the value as a member of the model instead.";
        }

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
        if (IsTokenValue(type))
        {
            memberType = null;
            return false;
        }

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

    // A token is a thing to print, not an object to read. The markdown flow hands Fluid the text it
    // renders as, so Source on an HtmlToken is nil there and renders blank, while the docx flow
    // prints raw markup, or the type name of the delegate an OpenXmlToken or a MutateToken holds.
    // A token type of the model's own is no different. Lockstep with the generator, whose
    // ShapeBuilder records no members for one, so no accessor exists either.
    static bool IsTokenValue(Type type) =>
        typeof(TokenValue).IsAssignableFrom(type);

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

    /// <summary>
    /// Resolves a member the way the source generator picks the one it emits an accessor for
    /// (ShapeBuilder walks the same order): the most derived type first, and within a type the
    /// first declared member whose name matches without regard to case.
    /// </summary>
    /// <remarks>
    /// Walked by hand because a case-insensitive <c>GetProperty</c> / <c>GetField</c> throws
    /// <see cref="AmbiguousMatchException"/> as soon as two members answer to the name, and both
    /// ways of getting there are legal C#: a <c>new</c> member hiding a base one of another type,
    /// and two members whose names differ only by case.
    /// </remarks>
    public static bool TryResolveMember(Type type, string name, [NotNullWhen(true)] out Type? memberType) =>
        TryResolveMember(type, name, out memberType, out _);

    /// <summary>
    /// Resolves a name against the model itself, which binds through its own accessors whatever
    /// else it implements.
    /// </summary>
    /// <remarks>
    /// A model that is a collection is the one case where that is narrower than reflection: the
    /// generator stops its base-chain walk at the first system base of such a type
    /// (<c>ShapeBuilder.BuildEntry</c>), so <c>Count</c> or <c>Keys</c> has no accessor and the
    /// generator reports a token naming one.
    /// </remarks>
    public static bool TryResolveRootMember(Type modelType, string name, [NotNullWhen(true)] out Type? memberType)
    {
        if (!TryResolveMember(modelType, name, out memberType, out var declaringType))
        {
            return false;
        }

        if (LiquidMembers.Classify(modelType) != LiquidKind.None &&
            IsSystemType(declaringType))
        {
            memberType = null;
            return false;
        }

        return true;
    }

    // Lockstep with the generator's ShapeBuilder.IsSystemType.
    static bool IsSystemType(Type? type)
    {
        var ns = type?.Namespace;
        return ns == "System" ||
               ns?.StartsWith("System.", StringComparison.Ordinal) == true;
    }

    static bool TryResolveMember(Type type, string name, [NotNullWhen(true)] out Type? memberType, out Type? declaringType)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
        for (var current = type; current != null; current = current.BaseType)
        {
            PropertyInfo? property = null;
            foreach (var candidate in current.GetProperties(flags))
            {
                // Readable in public, and not an indexer: the generator emits an accessor for
                // nothing else (ShapeBuilder.TryGetMemberType), so anything else would register
                // and then render blank.
                if (candidate is { GetMethod.IsPublic: true } &&
                    string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase) &&
                    candidate.GetIndexParameters().Length == 0 &&
                    (property == null || candidate.MetadataToken < property.MetadataToken))
                {
                    property = candidate;
                }
            }

            FieldInfo? field = null;
            foreach (var candidate in current.GetFields(flags))
            {
                if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase) &&
                    (field == null || candidate.MetadataToken < field.MetadataToken))
                {
                    field = candidate;
                }
            }

            if (field != null &&
                (property == null || IsDeclaredBefore(field, property)))
            {
                memberType = field.FieldType;
                declaringType = field.DeclaringType;
                return IsBindable(type, field, memberType);
            }

            if (property != null)
            {
                memberType = property.PropertyType;
                declaringType = property.DeclaringType;
                return !IsStaticInterfaceContract(property.GetMethod!) &&
                       IsBindable(type, property, memberType);
            }
        }

        memberType = null;
        declaringType = null;
        return false;
    }

    // Metadata keeps properties and fields in separate tables, each in declaration order, so the
    // two cannot be compared directly. An auto-property's backing field sits in the field table at
    // the property's position, which places it against a field. A property with no backing field
    // cannot be placed, and is taken to come first.
    static bool IsDeclaredBefore(FieldInfo field, PropertyInfo property)
    {
        var backing = property.DeclaringType!.GetField(
            $"<{property.Name}>k__BackingField",
            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
        return backing != null &&
               field.MetadataToken < backing.MetadataToken;
    }

    // A static abstract or virtual member of an interface is reachable through a type parameter
    // only, so there is no accessor the generator could emit for it.
    static bool IsStaticInterfaceContract(MethodInfo getter) =>
        getter.IsStatic &&
        (getter.IsAbstract || getter.IsVirtual);

    // The rest of what ShapeBuilder leaves out: a value that cannot become an object (a ref struct
    // such as a span, a pointer), and a member whose name a more derived type has reused for a
    // method, an event or a nested type, which is what the name then means through that type.
    static bool IsBindable(Type type, MemberInfo member, Type memberType)
    {
        if (memberType.IsByRefLike ||
            ContainsPointer(memberType))
        {
            return false;
        }

        for (var current = type; current != null && current != member.DeclaringType; current = current.BaseType)
        {
            var hiding = current.GetMember(
                member.Name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
            if (hiding.Length > 0)
            {
                return false;
            }
        }

        return true;
    }

    static bool ContainsPointer(Type type)
    {
        if (type.IsPointer ||
            type.IsFunctionPointer)
        {
            return true;
        }

        if (type.HasElementType)
        {
            // ReSharper disable once TailRecursiveCall
            return ContainsPointer(type.GetElementType()!);
        }

        foreach (var argument in type.GetGenericArguments())
        {
            if (ContainsPointer(argument))
            {
                return true;
            }
        }

        return false;
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
        return TryResolveRootMember(modelType, name, out rootType);
    }
}
