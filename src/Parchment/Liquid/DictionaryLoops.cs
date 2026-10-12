using System.Text.Encodings.Web;

/// <summary>
/// Makes the entry a <c>{% for kv in Lookup %}</c> loop hands out answer to <c>Key</c> and
/// <c>Value</c>, the names every validator binds the loop variable to.
/// </summary>
/// <remarks>
/// Fluid turns a dictionary into a <c>DictionaryValue</c>, and enumerating one yields a two-element
/// array per entry: <c>kv[0]</c>, <c>kv[1]</c>, <c>kv.first</c> and <c>kv.last</c> bind, while the
/// <c>KeyValuePair</c> the validators (and the accessors the source generator emits for it) describe
/// never reaches the template, so <c>kv.Key</c> resolved to nil and rendered blank. Both classes
/// here delegate everything to the Fluid value they wrap, so a dictionary and its entries keep
/// every liquid behaviour they had; only the two names are added. This sits on the value rather
/// than in the docx loop runner because the markdown flow iterates through Fluid's own
/// <c>for</c> statement, which only the value can reach.
/// </remarks>
static class DictionaryLoops
{
    // Set while Fluid is asked for its own conversion of the value the converter was handed, so
    // that the nested call falls through to Fluid instead of coming back here.
    [ThreadStatic]
    static bool converting;

    public static void Register(TemplateOptions options) =>
        options.ValueConverters
            .Add(value =>
            {
                if (converting ||
                    !IsDictionary(value))
                {
                    return null;
                }

                converting = true;
                FluidValue converted;
                try
                {
                    converted = FluidValue.Create(value, options);
                }
                finally
                {
                    converting = false;
                }

                if (converted is DictionaryValue dictionary)
                {
                    return new LoopDictionary(dictionary);
                }

                return converted;
            });

    static readonly ConcurrentDictionary<Type, bool> stringKeyed = new();

    // The two shapes Fluid makes a DictionaryValue of. Nothing else may take the nested call
    // above: Fluid converts the items of a list as it wraps the list, and a dictionary among them
    // would be converted while the flag is still set. A dictionary's values are converted on read.
    static bool IsDictionary(object value)
    {
        if (value is IDictionary)
        {
            return true;
        }

        if (value is string or not IEnumerable)
        {
            return false;
        }

        return stringKeyed.GetOrAdd(
            value.GetType(),
            static type => type
                .GetInterfaces()
                .Any(_ => _.IsGenericType &&
                          _.GetGenericTypeDefinition() == typeof(IDictionary<,>) &&
                          _.GetGenericArguments()[0] == typeof(string)));
    }

    static FluidValue Unwrap(FluidValue value)
    {
        if (value is LoopDictionary dictionary)
        {
            return dictionary.Inner;
        }

        if (value is Entry entry)
        {
            return entry.Pair;
        }

        return value;
    }

    sealed class LoopDictionary(DictionaryValue inner) :
        FluidValue
    {
        public DictionaryValue Inner => inner;

        public override FluidValues Type => inner.Type;

        // The order is the one Fluid reads the keys in, which is the dictionary's own.
        public override IEnumerable<FluidValue> Enumerate(TemplateContext context)
        {
            foreach (var item in inner.Enumerate(context))
            {
                if (item is ArrayValue { Values.Count: 2 } pair)
                {
                    yield return new Entry(pair);
                    continue;
                }

                yield return item;
            }
        }

        public override ValueTask<FluidValue> GetValueAsync(string name, TemplateContext context) =>
            inner.GetValueAsync(name, context);

        public override ValueTask<FluidValue> GetIndexAsync(FluidValue index, TemplateContext context) =>
            inner.GetIndexAsync(index, context);

        public override bool Contains(FluidValue value) =>
            inner.Contains(value);

        public override bool Equals(FluidValue other) =>
            inner.Equals(Unwrap(other));

        public override bool Equals(object? obj) =>
            obj is FluidValue other &&
            Equals(other);

        public override int GetHashCode() =>
            inner.GetHashCode();

        public override bool ToBooleanValue() =>
            inner.ToBooleanValue();

        public override decimal ToNumberValue() =>
            inner.ToNumberValue();

        public override string ToStringValue() =>
            inner.ToStringValue();

        public override object ToObjectValue() =>
            inner.ToObjectValue();

        public override ValueTask WriteToAsync(TextWriter writer, TextEncoder encoder, CultureInfo cultureInfo) =>
            inner.WriteToAsync(writer, encoder, cultureInfo);
    }

    sealed class Entry(ArrayValue pair) :
        FluidValue
    {
        public ArrayValue Pair => pair;

        public override FluidValues Type => pair.Type;

        // Without regard to case, as a model member is found; `size`, `first` and `last` stay
        // ordinal, being the array's own.
        public override ValueTask<FluidValue> GetValueAsync(string name, TemplateContext context)
        {
            if (TryGetNamed(name, out var named))
            {
                return named;
            }

            return pair.GetValueAsync(name, context);
        }

        // The validators read `kv['Value']` as `kv.Value`, and an array reads a bracketed string as
        // an index, which would answer with the key.
        public override ValueTask<FluidValue> GetIndexAsync(FluidValue index, TemplateContext context)
        {
            if (index.Type == FluidValues.String &&
                TryGetNamed(index.ToStringValue(), out var named))
            {
                return named;
            }

            return pair.GetIndexAsync(index, context);
        }

        bool TryGetNamed(string name, [NotNullWhen(true)] out FluidValue? value)
        {
            if (string.Equals(name, "Key", StringComparison.OrdinalIgnoreCase))
            {
                value = pair.Values[0];
                return true;
            }

            if (string.Equals(name, "Value", StringComparison.OrdinalIgnoreCase))
            {
                value = pair.Values[1];
                return true;
            }

            value = null;
            return false;
        }

        public override IEnumerable<FluidValue> Enumerate(TemplateContext context) =>
            pair.Enumerate(context);

        public override bool Contains(FluidValue value) =>
            pair.Contains(value);

        public override bool Equals(FluidValue other) =>
            pair.Equals(Unwrap(other));

        public override bool Equals(object? obj) =>
            obj is FluidValue other &&
            Equals(other);

        public override int GetHashCode() =>
            pair.GetHashCode();

        public override bool ToBooleanValue() =>
            pair.ToBooleanValue();

        public override decimal ToNumberValue() =>
            pair.ToNumberValue();

        public override string ToStringValue() =>
            pair.ToStringValue();

        public override object ToObjectValue() =>
            pair.ToObjectValue();

        public override ValueTask WriteToAsync(TextWriter writer, TextEncoder encoder, CultureInfo cultureInfo) =>
            pair.WriteToAsync(writer, encoder, cultureInfo);
    }
}
