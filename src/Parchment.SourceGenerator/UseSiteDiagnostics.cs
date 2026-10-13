/// <summary>
/// The compiler diagnostics a generated reference to a model member or type would raise at the use
/// site because of an attribute on it. The generated file names every bindable member, so a
/// deprecated or experimental one would otherwise warn, or fail the consumer's build, from source
/// the consumer cannot edit.
/// </summary>
static class UseSiteDiagnostics
{
    const string experimentalAttribute = "System.Diagnostics.CodeAnalysis.ExperimentalAttribute";

    /// <summary>
    /// Adds the ids to suppress for a reference to <paramref name="member"/>. False when the
    /// reference cannot be written at all (obsolete as an error, which no pragma silences).
    /// </summary>
    public static bool TryCollectMember(ISymbol member, List<string> ids)
    {
        if (!TryCollect(member, ids))
        {
            return false;
        }

        // The trimming and AOT attributes cannot go on a property, only on its accessors.
        if (member is IPropertySymbol property)
        {
            CollectRequires(property.GetMethod, ids);
            CollectRequires(property.SetMethod, ids);
        }

        return true;
    }

    /// <summary>
    /// Adds the ids to suppress for naming <paramref name="type"/>, which names its element type,
    /// its type arguments and the types it is nested in as well.
    /// </summary>
    public static bool TryCollectType(ITypeSymbol type, List<string> ids)
    {
        if (type is IArrayTypeSymbol array)
        {
            return TryCollectType(array.ElementType, ids);
        }

        if (type is not INamedTypeSymbol named)
        {
            return true;
        }

        if (!TryCollect(named, ids))
        {
            return false;
        }

        // An assembly or module can be experimental as a whole. Inside that assembly the compiler
        // stays silent and the pragma is inert.
        CollectExperimental(named.ContainingModule, ids);
        CollectExperimental(named.ContainingAssembly, ids);

        foreach (var argument in named.TypeArguments)
        {
            if (!TryCollectType(argument, ids))
            {
                return false;
            }
        }

        if (named.ContainingType == null)
        {
            return true;
        }

        return TryCollectType(named.ContainingType, ids);
    }

    static bool TryCollect(ISymbol symbol, List<string> ids)
    {
        foreach (var attribute in symbol.GetAttributes())
        {
            var name = attribute.AttributeClass?.ToDisplayString();
            if (name == "System.ObsoleteAttribute")
            {
                if (attribute.ConstructorArguments is [_, { Value: true }])
                {
                    return false;
                }

                // A custom DiagnosticId replaces the two compiler ids. One that is no identifier
                // cannot be named in a pragma, which leaves the member unreferenceable.
                if (TryGetNamedString(attribute, "DiagnosticId", out var custom))
                {
                    if (!SyntaxFacts.IsValidIdentifier(custom))
                    {
                        return false;
                    }

                    ids.Add(custom);
                    continue;
                }

                ids.Add("CS0612");
                ids.Add("CS0618");
                continue;
            }

            // Always an error, under the id the attribute carries.
            if (name == experimentalAttribute)
            {
                if (!TryGetExperimentalId(attribute, out var id))
                {
                    return false;
                }

                ids.Add(id);
                continue;
            }

            AddRequires(name, ids);
        }

        return true;
    }

    static void CollectExperimental(ISymbol? symbol, List<string> ids)
    {
        if (symbol == null)
        {
            return;
        }

        foreach (var attribute in symbol.GetAttributes())
        {
            if (attribute.AttributeClass?.ToDisplayString() == experimentalAttribute &&
                TryGetExperimentalId(attribute, out var id))
            {
                ids.Add(id);
            }
        }
    }

    static void CollectRequires(IMethodSymbol? accessor, List<string> ids)
    {
        if (accessor == null)
        {
            return;
        }

        foreach (var attribute in accessor.GetAttributes())
        {
            AddRequires(attribute.AttributeClass?.ToDisplayString(), ids);
        }
    }

    static void AddRequires(string? attributeName, List<string> ids)
    {
        switch (attributeName)
        {
            case "System.Diagnostics.CodeAnalysis.RequiresUnreferencedCodeAttribute":
                ids.Add("IL2026");
                break;
            case "System.Diagnostics.CodeAnalysis.RequiresAssemblyFilesAttribute":
                ids.Add("IL3002");
                break;
            case "System.Diagnostics.CodeAnalysis.RequiresDynamicCodeAttribute":
                ids.Add("IL3050");
                break;
        }
    }

    static bool TryGetExperimentalId(AttributeData attribute, [NotNullWhen(true)] out string? id)
    {
        if (attribute.ConstructorArguments is [{ Value: string value }] &&
            SyntaxFacts.IsValidIdentifier(value))
        {
            id = value;
            return true;
        }

        id = null;
        return false;
    }

    static bool TryGetNamedString(AttributeData attribute, string key, [NotNullWhen(true)] out string? value)
    {
        foreach (var named in attribute.NamedArguments)
        {
            if (named.Key == key &&
                named.Value.Value is string { Length: > 0 } text)
            {
                value = text;
                return true;
            }
        }

        value = null;
        return false;
    }
}
