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

    static TemplateOptions BuildOptions()
    {
        // No MaxSteps. Fluid counts a step for every statement it executes, so the steps a render
        // takes grow with the model rather than with the template: a limit caps how many rows a
        // loop can be handed, and a report that renders for a hundred rows fails for a thousand.
        // Running out is also reported as "The maximum level of recursion has been reached. Your
        // script must have a cyclic include statement", which names a cause the template does not
        // have. MaxRecursion is the limit that message belongs to, and it stays.
        var options = new TemplateOptions
        {
            MaxRecursion = 100
        };

        // A token names a model member without regard to case. Every validator resolves it that
        // way, and so do the per-model maps, which are keyed on the member's path. Fluid compares
        // ordinally unless told otherwise, and that left a token registration had accepted reading
        // nothing at render. It has to be set here, ahead of the first registered accessor: the
        // strategy chooses the comparer for its map inside Register, and a lookup uses whatever
        // the map was last built with.
        options.MemberAccessStrategy.IgnoreCasing = true;
        Filters.Register(options.Filters);
        // Route enum substitutions through Excelsior so inline `{{ Model.Status }}` tokens
        // render with the same humanization (Display attribute / source-gen switch /
        // ValueRenderer.ForEnums) that Excelsior applies to table cells. Returning a
        // StringValue short-circuits Fluid's default type dispatch (which would have called
        // Enum.ToString() and emitted the raw symbol name).
        options.ValueConverters
            .Add(static value =>
            {
                if (value is Enum e)
                {
                    return new StringValue(EnumRender.Render(e));
                }

                // Fluid turns a TimeSpan into a date counted from 1970 in the machine's time zone,
                // and prints a TimeOnly in the short pattern, which has no seconds. Both become
                // the invariant text of the value instead: a duration in the constant format, and
                // a time in the format an editable TimeOnly field defaults to.
                if (value is TimeSpan span)
                {
                    return new StringValue(span.ToString("c", CultureInfo.InvariantCulture));
                }

                if (value is TimeOnly time)
                {
                    return new StringValue(time.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                }

                return null;
            });
        DictionaryLoops.Register(options);

        return options;
    }

    static TemplateOptions BuildMarkdownOptions()
    {
        var options = BuildOptions();
        options.MemberAccessStrategy = Options.MemberAccessStrategy;
        options.ValueConverters
            .Add(static value =>
            {
                if (value is TokenValue token)
                {
                    // A TextToken is a plain string wearing a TokenValue coat, so it is escaped
                    // like any other bound value. A MarkdownToken or HtmlToken is markup the caller
                    // picked the type to say so — that declaration is what turns MarkdownEncoder
                    // off for them.
                    return new StringValue(TokenMarkdown.Render(token), encode: token is TextToken);
                }

                return null;
            });

        return options;
    }

    /// <summary>
    /// The context a render evaluates against, with the model held as an object.
    /// </summary>
    /// <remarks>
    /// Handing the model to <c>TemplateContext</c> as it is runs it through
    /// <c>FluidValue.Create</c>, which makes an <c>ArrayValue</c> or a <c>DictionaryValue</c> of a
    /// model that is itself a collection (<c>class Order : List&lt;Line&gt;</c>). Those answer only
    /// to <c>size</c>, <c>first</c> and <c>last</c>, or to a key, so every member the model declares
    /// rendered blank although all four validators accept it. Wrapped here, the root binds through
    /// its accessors like any other model. Only the root: the same type reached as a member or as a
    /// loop element is still converted by Fluid, which is what lets it be looped over, and
    /// <see cref="LiquidMembers"/> validates it as the array it is there.
    /// </remarks>
    public static TemplateContext CreateContext(object model, TemplateOptions options) =>
        new(new ObjectValue(model), options);

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

    /// <summary>
    /// Whether the source generator emitted <paramref name="modelType"/> as a model — its module
    /// initializer always emits the root type's accessor block, even for a member-less model.
    /// </summary>
    /// <remarks>
    /// Having accessors is not enough. The generator emits an accessor block for every type a model
    /// reaches, so a type it saw only as a member is in <see cref="registeredTypes"/> too, while
    /// the per-model maps — formats, string lists, tables, editable fields — are emitted for the
    /// model alone. Taken as a model, such a type would bind and then render with every attribute
    /// ignored. The attribute is what the generator keys on, so it is what tells the two apart;
    /// neither attribute is inherited, which leaves out a type merely derived from a model.
    /// </remarks>
    public static bool IsModelRegistered(Type modelType) =>
        registeredTypes.ContainsKey(modelType) &&
        (Attribute.IsDefined(modelType, typeof(ParchmentModelAttribute), false) ||
         Attribute.IsDefined(modelType, typeof(ParchmentBindableAttribute), false));

    /// <summary>
    /// Guards that <paramref name="modelType"/> arrived with pre-compiled accessors. There is no
    /// reflection fallback: every binding model is walked at compile time by the source generator,
    /// which emits one <c>DelegateAccessor</c> per member for every reachable type.
    /// </summary>
    public static void EnsureModelRegistered(Type modelType, string name)
    {
        if (IsModelRegistered(modelType))
        {
            return;
        }

        throw new ParchmentRegistrationException(
            name,
            NotGeneratedMessage(modelType));
    }

    public static string NotGeneratedMessage(Type modelType) =>
        $"Model '{modelType.FullName}' has no pre-compiled Parchment accessors. Mark the class with [ParchmentModel] (template found by convention) or [ParchmentBindable] (template supplied at runtime) and make it partial — the Parchment source generator emits the accessors at compile time.";

    /// <summary>
    /// Source-generator entry point (invoked via `Generated.GeneratedRegistration`).
    /// Registers pre-built accessors for a single type. That makes it bindable as a member or a
    /// loop element; <see cref="IsModelRegistered"/> decides whether it is also a model.
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
}
