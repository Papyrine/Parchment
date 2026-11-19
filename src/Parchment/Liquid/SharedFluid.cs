using DateTimeValue = Fluid.Values.DateTimeValue;
using StringValue = Fluid.Values.StringValue;

/// <summary>
/// Static singletons for Fluid. Fluid's parser, options, and filters are thread-safe and expensive
/// to construct; one instance per process is the documented recommendation.
/// </summary>
static class SharedFluid
{
    public static FluidParser Parser { get; } = new MarkdownFluidParser();

    // One for both option sets, and for every set AddFilter builds after them: the generated
    // accessors are registered once and reach both flows.
    static readonly PrecompiledMemberAccessStrategy memberAccess = new();

    // Fluid 3 options are immutable once built, so a filter added later means building both sets
    // again. Serialized so two filters added at once both survive.
    static readonly System.Threading.Lock filterSync = new();

    // Every markdown set ever built, so a render that began before AddFilter swapped the sets is
    // still recognised as the markdown flow by the filters it runs.
    static readonly ConditionalWeakTable<TemplateOptions, object> markdownSets = new();

    static volatile TemplateOptions options = BuildOptions();
    static volatile TemplateOptions markdownOptions = BuildMarkdownOptions();

    public static TemplateOptions Options => options;

    /// <summary>
    /// Options for the markdown flow, which needs <see cref="TokenValue"/> flattened to markdown
    /// source rather than left as an <c>ObjectValue</c>.
    /// </summary>
    /// <remarks>
    /// This cannot go on <see cref="Options"/>: the docx flow needs the TokenValue to survive as an
    /// ObjectValue so <c>ScopeTreeRunner.InterpretFluidValue</c> can dispatch it structurally.
    /// Without a converter the markdown flow falls through to Fluid's default object handling,
    /// which calls ToString() and writes the type name into the document. The one kind with nothing
    /// structural about it is a <see cref="TextToken"/>, and that one is converted for both flows,
    /// in <see cref="BuildOptions"/>. The MemberAccessStrategy is shared, so RegisterModel covers
    /// both sets of options.
    /// </remarks>
    public static TemplateOptions MarkdownOptions => markdownOptions;

    /// <summary>
    /// Whether <paramref name="templateOptions"/> is a set the markdown flow renders with.
    /// </summary>
    public static bool IsMarkdown(TemplateOptions templateOptions) =>
        markdownSets.TryGetValue(templateOptions, out _);

    static readonly ConcurrentDictionary<Type, bool> registeredTypes = new();

    static TemplateOptions BuildOptions() =>
        CreateBuilder(() => Options).Build();

    static TemplateOptionsBuilder CreateBuilder(Func<TemplateOptions> current)
    {
        // No MaxSteps. Fluid counts a step for every statement it executes, so the steps a render
        // takes grow with the model rather than with the template: a limit caps how many rows a
        // loop can be handed, and a report that renders for a hundred rows fails for a thousand.
        // Running out is also reported as "The maximum level of recursion has been reached. Your
        // script must have a cyclic include statement", which names a cause the template does not
        // have. MaxRecursion is the limit that message belongs to, and it stays.
        var builder = new TemplateOptionsBuilder()
            .WithMaxRecursion(100)
            // Fluid defaults this to TimeZoneInfo.Local and uses it to read a string or a unix
            // timestamp handed to the `date` filter, so the same template and model printed a
            // different time on a machine in another zone. 'now' and 'today' are read in it too,
            // so they print the machine's clock in UTC.
            .WithTimeZone(TimeZoneInfo.Utc)
            // A token names a model member without regard to case. Every validator resolves it that
            // way, and so do the per-model maps, which are keyed on the member's path. The strategy
            // compares member names that way itself.
            .WithMemberAccessStrategy(static () => memberAccess)
            // Variables, though, keep their exact spelling: `{% for color in Colors %}` binds
            // `color` and leaves the model's `Color` readable. Fluid 3 compares variable names with
            // this comparer and defaults it to ignoring case. The strategy is handed it too, and
            // does not use it.
            .WithModelNamesComparer(StringComparer.Ordinal)
            .ConfigureFilters(Filters.Register);

        // A DateTime of Kind Unspecified names no zone. Fluid keeps one as a PlainDateTimeValue,
        // read in the context's zone (UTC, above) when it is printed or formatted, but equal only
        // to another PlainDateTimeValue: `{% if Unspecified == Utc %}` was false for the same
        // wall clock. A DateTimeValue at offset zero prints the same text and compares as the
        // instant it names. Local is left to Fluid — it is an instant on the machine that built
        // it, which no fixed offset can stand in for.
        builder
            .AddValueConverter(static value =>
            {
                if (value is DateTime { Kind: DateTimeKind.Unspecified } dateTime)
                {
                    return DateTimeValue.Create(new DateTimeOffset(dateTime.Ticks, TimeSpan.Zero));
                }

                return null;
            });
        // Route enum substitutions through Excelsior so inline `{{ Model.Status }}` tokens
        // render with the same humanization (Display attribute / source-gen switch /
        // ValueRenderer.ForEnums) that Excelsior applies to table cells. Returning a
        // StringValue short-circuits Fluid's default type dispatch (which would have called
        // Enum.ToString() and emitted the raw symbol name).
        builder
            .AddValueConverter(static value =>
            {
                if (value is Enum e)
                {
                    return StringValue.Create(EnumRender.Render(e));
                }

                // A TextToken is a plain string wearing a TokenValue coat, so Fluid is handed the
                // string: what a filter receives, a condition compares and a substitution writes
                // is then what a string member would have produced, in either flow. Left as an
                // object it has no text form, and the docx flow wrote its type name. A null string
                // is nil, as it is from a string member — and returning null here would mean "not
                // converted".
                if (value is TextToken text)
                {
                    if (text.Value == null)
                    {
                        return NilValue.Instance;
                    }

                    return StringValue.Create(text.Value);
                }

                // Fluid turns a TimeSpan into a date counted from 1970 in the machine's time zone,
                // and prints a TimeOnly in the short pattern, which has no seconds. Both become
                // the invariant text of the value instead: a duration in the constant format, and
                // a time in the format an editable TimeOnly field defaults to.
                if (value is TimeSpan span)
                {
                    return StringValue.Create(span.ToString("c", CultureInfo.InvariantCulture));
                }

                if (value is Time time)
                {
                    return StringValue.Create(time.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                }

                return null;
            });
        DictionaryLoops.Register(builder, current);

        return builder;
    }

    static TemplateOptions BuildMarkdownOptions()
    {
        var builder = CreateBuilder(() => MarkdownOptions);
        builder
            .AddValueConverter(static value =>
            {
                if (value is TokenValue token)
                {
                    // A MarkdownToken, HtmlToken or OpenXmlToken is markup the caller picked the
                    // type to say so — that declaration is what turns MarkdownEncoder off for them.
                    // A TextToken makes no such claim and never gets this far: the converter
                    // BuildOptions added runs ahead of this one and has already made it a string,
                    // which is escaped like any other bound value.
                    return StringValue.Create(TokenMarkdown.Render(token), encode: false);
                }

                return null;
            });

        return RecordMarkdown(builder.Build());
    }

    static TemplateOptions RecordMarkdown(TemplateOptions built)
    {
        markdownSets.AddOrUpdate(built, built);
        return built;
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
        new(ObjectValue.Create(model), options);

    /// <summary>
    /// Registers a user filter against every option set.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Registering against a single set is a silent no-op in the other flow: Fluid treats an
    /// unknown filter as a pass-through, so the value renders unfiltered with nothing reported.
    /// </para>
    /// <para>
    /// Fluid 3 options cannot change once built, so both sets are built again with the filter and
    /// replace the ones in use. A render already under way keeps the set it started with. The new
    /// sets carry over everything else, the shared member access strategy included.
    /// </para>
    /// </remarks>
    public static void AddFilter(string name, FilterDelegate filter)
    {
        lock (filterSync)
        {
            options = options.ToBuilder().AddFilter(name, filter).Build();
            markdownOptions = RecordMarkdown(markdownOptions.ToBuilder().AddFilter(name, filter).Build());
        }
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
        IEnumerable<KeyValuePair<string, MemberAccessor>> accessors)
    {
        if (!registeredTypes.TryAdd(type, true))
        {
            return;
        }

        memberAccess.Register(type, accessors);
    }
}
