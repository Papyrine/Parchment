/// <summary>
/// Cache of <c>[EditableField]</c>-marked members reachable from a model type, keyed by dotted
/// path from the root model. Populated by the source generator's module initializers (and
/// consulted by <c>ParchmentExtractor</c>); render-time lookup is a single dictionary hit.
/// Model-shape rules (supported type, usable setter, no conflicting attributes, collection
/// constraints) are enforced at compile time as PARCH013–PARCH015 / PARCH022.
/// </summary>
sealed class EditableMap
{
    static ConcurrentDictionary<Type, EditableMap> precompiledCache = new();

    Dictionary<string, EditableEntry> entries;
    Dictionary<string, CollectionEntry> collections;

    EditableMap(Dictionary<string, EditableEntry> entries, Dictionary<string, CollectionEntry> collections)
    {
        this.entries = entries;
        this.collections = collections;
    }

    public static EditableMap Empty { get; } = new(
        new(StringComparer.OrdinalIgnoreCase),
        new(StringComparer.OrdinalIgnoreCase));

    public bool IsEmpty => entries.Count == 0 && collections.Count == 0;

    /// <summary>
    /// Whether any editable <em>collection</em> (repeating-section) member is registered. The loop
    /// renderer checks this before probing a loop's source against the collection map, so a template
    /// with no editable collections skips that lookup entirely.
    /// </summary>
    public bool HasCollections => collections.Count > 0;

    /// <summary>
    /// Whether any reachable editable member — directly or inside an editable collection's element
    /// type — is rich-text HTML. Such a field lets the user apply bullets and numbering, which Word
    /// can only do against a list definition that already exists (see
    /// <c>WordNumbering.EnsureListDefinitions</c>), so registration seeds one when this is true.
    /// </summary>
    public bool HasHtmlField =>
        entries.Values.Any(_ => _.Kind == EditableFieldKind.Html) ||
        collections.Values.Any(_ => _.ElementMap.HasHtmlField);

    public IReadOnlyCollection<EditableEntry> Entries => entries.Values;

    public IReadOnlyCollection<CollectionEntry> Collections => collections.Values;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out EditableEntry? entry) =>
        entries.TryGetValue(dottedPath, out entry);

    public bool TryGetCollection(string dottedPath, [NotNullWhen(true)] out CollectionEntry? entry) =>
        collections.TryGetValue(dottedPath, out entry);

    /// <summary>
    /// Projects this element-type map to a per-item render map: each entry is re-keyed under the loop
    /// variable (e.g. <c>b.Year</c>) with its getter bound to <paramref name="item"/>. The entry's
    /// <see cref="EditableEntry.DottedPath"/> — the control tag — stays item-relative (e.g. <c>Year</c>),
    /// which is exactly what extraction reads inside each repeated section item.
    /// </summary>
    internal EditableMap ScopedToItem(string loopVariable, object? item)
    {
        var scoped = new Dictionary<string, EditableEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries.Values)
        {
            scoped[$"{loopVariable}.{entry.DottedPath}"] = entry with
            {
                Getter = _ => item == null ? null : entry.Getter(item),
                CanReach = static _ => true
            };
        }

        return new(scoped, new(StringComparer.OrdinalIgnoreCase));
    }

    // No reflection fallback: the source generator walks the model graph at compile time,
    // enforces the shape rules as PARCH013–PARCH015 / PARCH022, and registers the entries via
    // GeneratedRegistration. A model with no [EditableField] members has no registration, which
    // is the empty map.
    public static EditableMap Build(Type modelType) =>
        precompiledCache.GetValueOrDefault(modelType, Empty);

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<EditableFieldMapEntry> entries)
    {
        var dict = BuildEntryDict(entries);
        // Merge, so RegisterEditable and RegisterEditableCollections (emitted as separate calls) each
        // contribute their half regardless of order.
        precompiledCache.AddOrUpdate(
            modelType,
            _ => new(dict, EmptyCollections()),
            (_, existing) => new(dict, existing.collections));
    }

    internal static void RegisterPrecompiledCollections(Type modelType, IEnumerable<CollectionFieldMapEntry> entries)
    {
        var dict = new Dictionary<string, CollectionEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = new(
                entry.DottedPath,
                entry.ElementType,
                entry.Setter,
                entry.CanReach,
                entry.ElementFactory,
                new(BuildEntryDict(entry.ElementFields), EmptyCollections()),
                entry.IsArray);
        }

        precompiledCache.AddOrUpdate(
            modelType,
            _ => new(new(StringComparer.OrdinalIgnoreCase), dict),
            (_, existing) => new(existing.entries, dict));
    }

    static Dictionary<string, EditableEntry> BuildEntryDict(IEnumerable<EditableFieldMapEntry> entries)
    {
        var dict = new Dictionary<string, EditableEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
        {
            dict[entry.DottedPath] = new(
                entry.DottedPath,
                entry.Kind,
                entry.ClrType,
                entry.IsNullable,
                entry.Getter,
                entry.Setter,
                entry.CanReach,
                entry.MultiLine,
                entry.DateFormat);
        }

        return dict;
    }

    static Dictionary<string, CollectionEntry> EmptyCollections() =>
        new(StringComparer.OrdinalIgnoreCase);
}
