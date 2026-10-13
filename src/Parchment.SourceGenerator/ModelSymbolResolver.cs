// The collection types are recognised by name, arity and namespace read straight off the symbol.
// Formatting a display string to compare allocates per interface of every type in every model, on
// every compilation, since the shape is rebuilt each time.
static class ModelSymbolResolver
{
    public static bool TryGetElementType(ITypeSymbol type, [NotNullWhen(true)] out ITypeSymbol? elementType)
    {
        if (type is IArrayTypeSymbol array)
        {
            elementType = array.ElementType;
            return true;
        }

        foreach (var i in type.AllInterfaces)
        {
            if (IsGenericEnumerable(i))
            {
                elementType = i.TypeArguments[0];
                return true;
            }
        }

        if (type is INamedTypeSymbol named &&
            IsGenericEnumerable(named))
        {
            elementType = named.TypeArguments[0];
            return true;
        }

        elementType = null;
        return false;
    }

    public static bool IsKeyValuePair(INamedTypeSymbol type) =>
        IsGenericCollection(type, "KeyValuePair", 2);

    public static bool IsDictionary(ITypeSymbol type)
    {
        if (type is not INamedTypeSymbol {ContainingType: null} named)
        {
            return false;
        }

        if (named.Arity == 0)
        {
            return named.Name == "IDictionary" &&
                   IsSystemCollections(named.ContainingNamespace);
        }

        return IsGenericCollection(named, "IDictionary", 2) ||
               IsGenericCollection(named, "IReadOnlyDictionary", 2);
    }

    static bool IsGenericEnumerable(INamedTypeSymbol type) =>
        IsGenericCollection(type, "IEnumerable", 1);

    // A type nested in another, or declared in a namespace that only ends the same way, is not
    // the System.Collections.Generic one.
    static bool IsGenericCollection(INamedTypeSymbol type, string name, int arity) =>
        type.Arity == arity &&
        type.ContainingType == null &&
        type.Name == name &&
        type.ContainingNamespace is {Name: "Generic"} generic &&
        IsSystemCollections(generic.ContainingNamespace);

    static bool IsSystemCollections(INamespaceSymbol? ns) =>
        ns is {Name: "Collections", ContainingNamespace: {Name: "System", ContainingNamespace.IsGlobalNamespace: true}};
}
