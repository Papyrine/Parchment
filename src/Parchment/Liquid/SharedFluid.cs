using StringValue = Fluid.Values.StringValue;

/// <summary>
/// Static singletons for Fluid. Fluid's parser, options, and filters are thread-safe and expensive
/// to construct; one instance per process is the documented recommendation.
/// </summary>
static class SharedFluid
{
    public static FluidParser Parser { get; } = new();

    public static TemplateOptions Options { get; } = BuildOptions();

    /// <summary>
    /// Options for the markdown flow, which needs <see cref="TokenValue"/> flattened to markdown
    /// source rather than left as an <c>ObjectValue</c>.
    /// </summary>
    /// <remarks>
    /// This cannot go on <see cref="Options"/>: the docx flow needs the TokenValue to survive as an
    /// ObjectValue so <c>ScopeTreeRunner.InterpretFluidValue</c> can dispatch it structurally.
    /// Without a converter the markdown flow falls through to Fluid's default object handling,
    /// which calls ToString() and writes the type name into the document. The MemberAccessStrategy
    /// is shared, so RegisterModel covers both sets of options.
    /// </remarks>
    public static TemplateOptions MarkdownOptions { get; } = BuildMarkdownOptions();

    static readonly ConcurrentDictionary<Type, bool> registeredTypes = new();

    static readonly MethodInfo registerGenericMethod = typeof(MemberAccessStrategyExtensions)
        .GetMethods(BindingFlags.Public | BindingFlags.Static)
        .First(_ => _ is
                    {
                        Name: "Register",
                        IsGenericMethodDefinition: true
                    } &&
                    _.GetGenericArguments().Length == 1 &&
                    _.GetParameters().Length == 1);

    static TemplateOptions BuildOptions()
    {
        var options = new TemplateOptions
        {
            MaxSteps = 10_000,
            MaxRecursion = 100
        };
        Filters.Register(options.Filters);
        // Route enum substitutions through Excelsior so inline `{{ Model.Status }}` tokens
        // render with the same humanization (Display attribute / source-gen switch /
        // ValueRenderer.ForEnums) that Excelsior applies to table cells. Returning a
        // StringValue short-circuits Fluid's default type dispatch (which would have called
        // Enum.ToString() and emitted the raw symbol name).
        options.ValueConverters.Add(static value =>
        {
            if (value is Enum e)
            {
                return new StringValue(EnumRender.Render(e));
            }

            return null;
        });

        return options;
    }

    static TemplateOptions BuildMarkdownOptions()
    {
        var options = BuildOptions();
        options.MemberAccessStrategy = Options.MemberAccessStrategy;
        options.ValueConverters.Add(static value =>
        {
            if (value is TokenValue token)
            {
                return new StringValue(TokenMarkdown.Render(token));
            }

            return null;
        });

        return options;
    }

    /// <summary>
    /// Registers a user filter against every option set.
    /// </summary>
    /// <remarks>
    /// <see cref="Options"/> and <see cref="MarkdownOptions"/> hold separate FilterCollections.
    /// Fluid exposes <c>TemplateOptions.Filters</c> as get-only, so the two cannot be pointed at one
    /// instance the way <c>MemberAccessStrategy</c> is. Registering against a single set is a silent
    /// no-op in the other flow: Fluid treats an unknown filter as a pass-through, so the value
    /// renders unfiltered with nothing reported.
    /// </remarks>
    public static void AddFilter(string name, FilterDelegate filter)
    {
        Options.Filters.AddFilter(name, filter);
        MarkdownOptions.Filters.AddFilter(name, filter);
    }

    public static void RegisterModel(Type modelType) =>
        RegisterTypeGraph(modelType);

    /// <summary>
    /// Source-generator entry point (invoked via `Generated.GeneratedRegistration`).
    /// Registers pre-built accessors for a single type and marks it as visited so the reflection
    /// walk in <see cref="RegisterTypeGraph"/> short-circuits when the same type is later
    /// encountered through <see cref="RegisterModel"/>.
    /// </summary>
    internal static void RegisterPrecompiledAccessors(
        Type type,
        IEnumerable<KeyValuePair<string, IMemberAccessor>> accessors)
    {
        if (!registeredTypes.TryAdd(type, true))
        {
            return;
        }

        Options.MemberAccessStrategy.Register(type, accessors);
    }

    static void RegisterTypeGraph(Type? type)
    {
        if (type == null ||
            !ShouldRegister(type))
        {
            return;
        }

        if (!registeredTypes.TryAdd(type, true))
        {
            return;
        }

        // Fluid's MemberAccessStrategyExtensions.Register<T>() walks instance properties only
        // (and not fields, not static members). Use it for instance properties, then layer
        // explicit DelegateAccessors on top for everything else: static properties, instance
        // fields, static fields. PropertyInfo.GetValue / FieldInfo.GetValue ignore the obj
        // parameter for static members, so a single lambda shape works for both.
        var generic = registerGenericMethod.MakeGenericMethod(type);
        generic.Invoke(null, [Options.MemberAccessStrategy]);

        var staticProperties = type.GetProperties(BindingFlags.Public | BindingFlags.Static);
        var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static);

        var extra = new List<KeyValuePair<string, IMemberAccessor>>(staticProperties.Length + fields.Length);
        foreach (var property in staticProperties)
        {
            if (property.GetIndexParameters().Length > 0)
            {
                continue;
            }

            var captured = property;
            extra.Add(new(property.Name, new DelegateAccessor((_, _) => captured.GetValue(null))));
        }

        foreach (var field in fields)
        {
            var captured = field;
            extra.Add(new(field.Name, new DelegateAccessor((instance, _) => captured.GetValue(instance))));
        }

        if (extra.Count > 0)
        {
            Options.MemberAccessStrategy.Register(type, extra);
        }

        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
        {
            if (property.GetIndexParameters().Length > 0)
            {
                continue;
            }

            RegisterTypeGraph(Unwrap(property.PropertyType));
        }

        foreach (var field in fields)
        {
            RegisterTypeGraph(Unwrap(field.FieldType));
        }
    }

    static Type Unwrap(Type type)
    {
        if (type.IsArray)
        {
            return type.GetElementType() ?? type;
        }

        if (type.IsGenericType &&
            type.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            return type.GetGenericArguments()[0];
        }

        // Find the IEnumerable<T> interface and use its T. For Dictionary<K, V> this resolves to
        // KeyValuePair<K, V> — taking GetGenericArguments()[0] on the type itself would have
        // returned K instead, silently dropping V's members from the type-graph walk.
        foreach (var i in type.GetInterfaces())
        {
            if (i.IsGenericType &&
                i.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            {
                return i.GetGenericArguments()[0];
            }
        }

        if (type.IsGenericType &&
            type.GetGenericTypeDefinition() == typeof(IEnumerable<>))
        {
            return type.GetGenericArguments()[0];
        }

        return type;
    }

    static bool ShouldRegister(Type type)
    {
        if (type.IsPrimitive ||
            type == typeof(string) ||
            type == typeof(decimal) ||
            type == typeof(DateTime) ||
            type == typeof(DateTimeOffset) ||
            type == typeof(TimeSpan) ||
            type == typeof(Date) ||
            type == typeof(Time) ||
            type == typeof(Guid))
        {
            return false;
        }

        if (type.IsEnum)
        {
            return false;
        }

        // KeyValuePair<K, V> is in the System namespace but it's the iteration element for every
        // IDictionary<K, V> reached from the model. Registering it lets `{{ kv.Key }}` and
        // `{{ kv.Value }}` resolve, and the property-type recursion then visits V (and K) so
        // user-type values are reachable.
        if (type.IsGenericType &&
            type.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
        {
            return true;
        }

        if (type.Namespace?.StartsWith("System", StringComparison.Ordinal) == true)
        {
            return false;
        }

        return type.IsClass ||
               type is
               {
                   IsValueType: true,
                   IsPrimitive: false,
                   IsEnum: false
               };
    }
}
