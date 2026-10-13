/// <summary>
/// The members liquid itself provides on a value that Fluid does not hand to a template as an
/// object: a string, a dictionary, or anything else enumerable, which becomes an array.
/// </summary>
/// <remarks>
/// <c>FluidValue.Create</c> picks the wrapper, and the wrapper answers member access — the
/// registered accessors and the CLR members are never consulted. <c>ArrayValue</c> and
/// <c>StringValue</c> answer <c>size</c>, <c>first</c> and <c>last</c>; <c>DictionaryValue</c>
/// answers <c>size</c> and treats any other name as a key. All three compare the name ordinally,
/// so <c>Size</c> binds nothing. Everything else — <c>Count</c>, <c>Length</c>, a property of a
/// model's own collection class — resolves to nil and renders blank, which is why validation
/// resolves these types here instead of by reflection. The source generator mirrors this in
/// <c>ShapeResolver</c>, keyed off the <c>LiquidKind</c> its <c>ShapeBuilder</c> records.
/// <para>
/// An enum is the same case by another route: the value converter in <see cref="SharedFluid"/>
/// hands it over as the text it prints as. Reflection lists its constants, and the
/// <c>value__</c> it stores its number in, as public fields, and none of them bind — so
/// <c>{{ Status.Final }}</c> is blank and <c>{% if Status == Status.Final %}</c> is false for every
/// model. An enum resolves no member at all.
/// </para>
/// </remarks>
[RequiresUnreferencedCode(Unsupported.Trimming)]
static class LiquidMembers
{
    public static LiquidKind Classify(Type type)
    {
        if (type == typeof(string))
        {
            return LiquidKind.String;
        }

        if (EnumType(type).IsEnum)
        {
            return LiquidKind.Enum;
        }

        if (IsDictionary(type))
        {
            return LiquidKind.Dictionary;
        }

        if (typeof(IEnumerable).IsAssignableFrom(type))
        {
            return LiquidKind.Array;
        }

        return LiquidKind.None;
    }

    public static bool TryResolve(LiquidKind kind, Type type, string name, [NotNullWhen(true)] out Type? memberType)
    {
        memberType = null;
        if (kind == LiquidKind.Enum)
        {
            return false;
        }

        if (name == "size")
        {
            memberType = typeof(int);
            return true;
        }

        if (kind == LiquidKind.Dictionary ||
            name is not ("first" or "last"))
        {
            return false;
        }

        if (kind == LiquidKind.String)
        {
            memberType = typeof(string);
            return true;
        }

        if (!ModelValidator.TryResolveElementType(type, out memberType))
        {
            // A non-generic enumerable: the element exists but nothing is known about it.
            memberType = typeof(object);
        }

        return true;
    }

    public static string MissingMessage(LiquidKind kind, Type type, string name)
    {
        if (kind == LiquidKind.Enum)
        {
            return $"Member '{name}' of '{EnumType(type).Name}' is not available to a template. " +
                   "Liquid sees an enum as the text it renders as, so its constants cannot be named: compare against that text instead.";
        }

        var prefix = $"Member '{name}' of '{type.Name}' is not available to a template. ";
        if (kind == LiquidKind.Dictionary)
        {
            return prefix + "Liquid sees a dictionary, of which only 'size' can be validated at registration.";
        }

        if (kind == LiquidKind.String)
        {
            return prefix + "Liquid sees a string, which has only 'size', 'first' and 'last'.";
        }

        return prefix + "Liquid sees a collection as an array, which has only 'size', 'first' and 'last'.";
    }

    // A nullable enum is boxed on its way to Fluid, so what arrives is the enum or nothing.
    static Type EnumType(Type type) =>
        Nullable.GetUnderlyingType(type) ?? type;

    // Fluid makes a DictionaryValue of a non-generic IDictionary or an IDictionary<string, V>. The
    // read-only interface is included because the declared type is all there is to go on here and
    // its usual implementations are both.
    static bool IsDictionary(Type type)
    {
        if (typeof(IDictionary).IsAssignableFrom(type) ||
            IsGenericDictionary(type))
        {
            return true;
        }

        foreach (var i in type.GetInterfaces())
        {
            if (IsGenericDictionary(i))
            {
                return true;
            }
        }

        return false;
    }

    static bool IsGenericDictionary(Type type)
    {
        if (!type.IsGenericType)
        {
            return false;
        }

        var definition = type.GetGenericTypeDefinition();
        return definition == typeof(IDictionary<,>) ||
               definition == typeof(IReadOnlyDictionary<,>);
    }
}

enum LiquidKind
{
    None,
    String,
    Array,
    Dictionary,
    Enum
}
