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
        options.ValueConverters
            .Add(static value =>
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
    /// Whether the source generator registered <paramref name="modelType"/> — its module
    /// initializer always emits the root type's accessor block, even for a member-less model.
    /// </summary>
    public static bool IsModelRegistered(Type modelType) =>
        registeredTypes.ContainsKey(modelType);

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
    /// Registers pre-built accessors for a single type and marks it known to
    /// <see cref="EnsureModelRegistered"/>.
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
