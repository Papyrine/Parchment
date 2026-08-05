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
            if (i.IsGenericType &&
                i.ConstructedFrom.ToDisplayString() == "System.Collections.Generic.IEnumerable<T>")
            {
                elementType = i.TypeArguments[0];
                return true;
            }
        }

        if (type is INamedTypeSymbol { IsGenericType: true } named &&
            named.ConstructedFrom.ToDisplayString() == "System.Collections.Generic.IEnumerable<T>")
        {
            elementType = named.TypeArguments[0];
            return true;
        }

        elementType = null;
        return false;
    }
}
