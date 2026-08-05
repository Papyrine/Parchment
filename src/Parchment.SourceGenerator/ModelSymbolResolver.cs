static class ModelSymbolResolver
{
    public static ITypeSymbol? TryGetElementType(ITypeSymbol type)
    {
        if (type is IArrayTypeSymbol array)
        {
            return array.ElementType;
        }

        foreach (var i in type.AllInterfaces)
        {
            if (i.IsGenericType &&
                i.ConstructedFrom.ToDisplayString() == "System.Collections.Generic.IEnumerable<T>")
            {
                return i.TypeArguments[0];
            }
        }

        if (type is INamedTypeSymbol { IsGenericType: true } named &&
            named.ConstructedFrom.ToDisplayString() == "System.Collections.Generic.IEnumerable<T>")
        {
            return named.TypeArguments[0];
        }

        return null;
    }
}
