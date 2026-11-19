/// <summary>
/// The member access strategy both option sets share: the accessors the source generator emitted,
/// and nothing else.
/// </summary>
/// <remarks>
/// <para>
/// Fluid 3 freezes its <c>DefaultMemberAccessStrategy</c> once options are built, and the freeze
/// cannot be lifted from outside Fluid. Accessors arrive later than that: each assembly holding
/// models registers its own from a module initializer, whenever it loads. Deriving from
/// <see cref="MemberAccessStrategy"/> directly is what Fluid leaves open for a strategy that stays
/// writable.
/// </para>
/// <para>
/// The lookup is the one Fluid 2 made, which every validator was written against: the exact type,
/// then its base types, then its interfaces, with <c>*</c> answering for any member, and member
/// names compared without regard to case. Fluid 3's default strategy also reads any public member
/// of an unregistered type through reflection; this one does not, so a member no model declares
/// still renders blank rather than whatever the runtime type happens to expose.
/// </para>
/// </remarks>
sealed class PrecompiledMemberAccessStrategy :
    MemberAccessStrategy
{
    readonly System.Threading.Lock sync = new();

    // Replaced on every registration and never changed after, so a render reads it without a lock
    // and it doubles as the cache token.
    volatile Dictionary<Key, MemberAccessor> map = new(KeyComparer.Instance);
    volatile bool hasWildcard;

    readonly record struct Key(Type Type, string Name);

    // Fluid remembers the accessor a call site resolved and asks again only once this changes.
    protected override object AccessorCacheToken => map;

    // The comparer Fluid hands over is the options' model names comparer, which also governs
    // variable names. Members are matched without regard to case whatever it is.
    public override MemberAccessor? GetAccessor(Type type, string name, StringComparer stringComparer)
    {
        var current = map;
        if (TryGetAccessor(current, type, name, out var accessor))
        {
            return accessor;
        }

        var baseType = type.BaseType;
        while (baseType != typeof(object) &&
               baseType != null)
        {
            if (TryGetAccessor(current, baseType, name, out accessor))
            {
                return accessor;
            }

            baseType = baseType.BaseType;
        }

        foreach (var interfaceType in Interfaces(type))
        {
            if (TryGetAccessor(current, interfaceType, name, out accessor))
            {
                return accessor;
            }
        }

        return null;
    }

    [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "Only registered interfaces are looked up, and every one was registered from code the generator emitted against the type.")]
    static Type[] Interfaces(Type type) =>
        type.GetInterfaces();

    bool TryGetAccessor(Dictionary<Key, MemberAccessor> current, Type type, string name, [NotNullWhen(true)] out MemberAccessor? accessor) =>
        current.TryGetValue(new(type, name), out accessor) ||
        (hasWildcard && current.TryGetValue(new(type, "*"), out accessor));

    public override void Register(Type type, string name, MemberAccessor accessor) =>
        Register(type, [new(name, accessor)]);

    public void Register(Type type, IEnumerable<KeyValuePair<string, MemberAccessor>> accessors)
    {
        lock (sync)
        {
            var updated = new Dictionary<Key, MemberAccessor>(map, KeyComparer.Instance);
            var wildcard = hasWildcard;
            foreach (var (name, accessor) in accessors)
            {
                wildcard |= name == "*";
                updated[new(type, name)] = accessor;
            }

            hasWildcard = wildcard;
            map = updated;
        }
    }

    sealed class KeyComparer :
        IEqualityComparer<Key>
    {
        public static KeyComparer Instance { get; } = new();

        public bool Equals(Key x, Key y) =>
            x.Type == y.Type &&
            string.Equals(x.Name, y.Name, StringComparison.OrdinalIgnoreCase);

        public int GetHashCode(Key key) =>
            HashCode.Combine(key.Type, StringComparer.OrdinalIgnoreCase.GetHashCode(key.Name));
    }
}
