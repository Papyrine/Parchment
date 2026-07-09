/// <summary>
/// Cache of <c>[EditableField]</c>-marked members reachable from a model type, keyed by dotted
/// path from the root model. Built once at template registration time (and consulted by
/// <c>ParchmentExtractor</c>); render-time lookup is a single dictionary hit. Model-shape rules
/// (supported type, usable setter, no conflicting attributes) are enforced here so misuse fails
/// at registration — the source generator mirrors the same rules as PARCH013–PARCH015.
/// </summary>
sealed class EditableMap
{
    static ConcurrentDictionary<Type, EditableMap> precompiledCache = new();

    Dictionary<string, EditableEntry> entries;

    EditableMap(Dictionary<string, EditableEntry> entries) =>
        this.entries = entries;

    public static EditableMap Empty { get; } = new(new(StringComparer.OrdinalIgnoreCase));

    public bool IsEmpty => entries.Count == 0;

    public IReadOnlyCollection<EditableEntry> Entries => entries.Values;

    public bool TryGet(string dottedPath, [NotNullWhen(true)] out EditableEntry? entry) =>
        entries.TryGetValue(dottedPath, out entry);

    public static EditableMap Build(Type modelType, string templateName)
    {
        if (precompiledCache.TryGetValue(modelType, out var cached))
        {
            return cached;
        }

        var entries = new Dictionary<string, EditableEntry>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<Type>
        {
            modelType
        };
        // NullabilityInfoContext is not thread-safe — one per build.
        var nullability = new NullabilityInfoContext();
        WalkType(modelType, [], static root => root, entries, visited, nullability, templateName);
        return new(entries);
    }

    internal static void RegisterPrecompiled(Type modelType, IEnumerable<EditableFieldMapEntry> entries)
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

        precompiledCache[modelType] = new(dict);
    }

    static void WalkType(
        Type type,
        List<string> pathSegments,
        Func<object, object?> parentGetter,
        Dictionary<string, EditableEntry> entries,
        HashSet<Type> visited,
        NullabilityInfoContext nullability,
        string templateName)
    {
        foreach (var (name, memberType, memberGetter, member) in FormatMap.EnumerateMembers(type))
        {
            var nextSegments = new List<string>(pathSegments)
            {
                name
            };
            var getter = ChainGetter(parentGetter, memberGetter);

            var attribute = member.GetCustomAttribute<EditableFieldAttribute>(true);
            if (attribute != null)
            {
                var dottedPath = string.Join('.', nextSegments);
                entries[dottedPath] = BuildEntry(
                    dottedPath,
                    type,
                    member,
                    memberType,
                    attribute,
                    parentGetter,
                    getter,
                    nullability,
                    templateName);
                continue;
            }

            var memberUnderlying = Nullable.GetUnderlyingType(memberType) ?? memberType;
            if (!ShouldDescend(memberUnderlying))
            {
                continue;
            }

            if (!visited.Add(memberUnderlying))
            {
                continue;
            }

            WalkType(memberUnderlying, nextSegments, getter, entries, visited, nullability, templateName);
            visited.Remove(memberUnderlying);
        }
    }

    static EditableEntry BuildEntry(
        string dottedPath,
        Type owner,
        MemberInfo member,
        Type memberType,
        EditableFieldAttribute attribute,
        Func<object, object?> parentGetter,
        Func<object, object?> getter,
        NullabilityInfoContext nullability,
        string templateName)
    {
        var isHtml = GuardConflictingAttributes(owner, member, templateName);

        var underlying = Nullable.GetUnderlyingType(memberType);
        var effective = underlying ?? memberType;

        EditableFieldKind kind;
        if (isHtml)
        {
            // [Html] + [EditableField] renders an editable rich-content block that extracts back
            // to an HTML string — so the member must be a string.
            if (effective != typeof(string))
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] member '{owner.Name}.{member.Name}' is marked [Html] but is '{memberType.Name}', not string. Editable HTML round-trips a string value.");
            }

            kind = EditableFieldKind.Html;
        }
        else
        {
            var mapped = MapKind(effective);
            if (mapped == null)
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] member '{owner.Name}.{member.Name}' has unsupported type '{memberType.Name}'. Supported: string, bool, DateOnly, DateTime, DateTimeOffset, TimeOnly, enums, and numeric types (nullable variants except bool?).");
            }

            if (mapped == EditableFieldKind.Checkbox &&
                underlying != null)
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] member '{owner.Name}.{member.Name}' is bool? — a checkbox cannot represent null. Use a non-nullable bool.");
            }

            kind = mapped.Value;
        }

        var setter = BuildSetter(owner, member, parentGetter, templateName);
        var isNullable = underlying != null || IsNullableReference(member, nullability);

        return new(
            dottedPath,
            kind,
            effective,
            isNullable,
            getter,
            setter,
            root => parentGetter(root) != null,
            attribute.MultiLine,
            attribute.DateFormat);
    }

    /// <summary>
    /// Rejects attribute combinations that can't coexist with <c>[EditableField]</c> and reports
    /// whether the member opts into editable HTML. <c>[Html]</c> (or <c>[StringSyntax("html")]</c>)
    /// is allowed — it selects the rich-content <see cref="EditableFieldKind.Html"/> control.
    /// <c>[ExcelsiorTable]</c> is always a conflict; <c>[Markdown]</c> is rejected because editable
    /// round-trip is HTML-only (extraction has no OpenXML→Markdown serializer).
    /// </summary>
    static bool GuardConflictingAttributes(Type owner, MemberInfo member, string templateName)
    {
        var isHtml = false;
        foreach (var attribute in member.GetCustomAttributes(true))
        {
            var name = attribute.GetType().Name;
            if (attribute is ExcelsiorTableAttribute)
            {
                throw Conflict(owner, member, templateName, "ExcelsiorTable");
            }

            if (name == "MarkdownAttribute")
            {
                throw MarkdownNotSupported(owner, member, templateName);
            }

            if (name == "HtmlAttribute")
            {
                isHtml = true;
            }
            else if (attribute is StringSyntaxAttribute syntax)
            {
                switch (syntax.Syntax.ToLowerInvariant())
                {
                    case "markdown":
                        throw MarkdownNotSupported(owner, member, templateName);
                    case "html":
                        isHtml = true;
                        break;
                }
            }
        }

        return isHtml;
    }

    static ParchmentRegistrationException Conflict(Type owner, MemberInfo member, string templateName, string other) =>
        new(
            templateName,
            $"Member '{owner.Name}.{member.Name}': [EditableField] cannot be combined with [{other}] — an editable field is plain typed content, not rendered markup.");

    static ParchmentRegistrationException MarkdownNotSupported(Type owner, MemberInfo member, string templateName) =>
        new(
            templateName,
            $"Member '{owner.Name}.{member.Name}': [EditableField] supports editable rich text via [Html] only — [Markdown] round-trip is not supported. Use [Html], or drop [EditableField] to render read-only Markdown.");

    static Action<object, object?> BuildSetter(Type owner, MemberInfo member, Func<object, object?> parentGetter, string templateName)
    {
        // The setter receives the ROOT model; the member lives on the object at the end of the
        // parent path. Callers (ExtractResult.ApplyTo) validate reachability via CanReach before
        // invoking, so the parent is non-null here.
        if (member is PropertyInfo property)
        {
            var set = property.SetMethod;
            if (set is not { IsPublic: true } || IsInitOnly(set))
            {
                throw new ParchmentRegistrationException(
                    templateName,
                    $"[EditableField] member '{owner.Name}.{member.Name}' has no public non-init setter. Extraction writes values back onto the model, so the member must be settable.");
            }

            return (root, value) => property.SetValue(parentGetter(root), value);
        }

        var field = (FieldInfo)member;
        if (field.IsInitOnly)
        {
            throw new ParchmentRegistrationException(
                templateName,
                $"[EditableField] member '{owner.Name}.{member.Name}' is a readonly field. Extraction writes values back onto the model, so the member must be settable.");
        }

        return (root, value) => field.SetValue(parentGetter(root), value);
    }

    static bool IsInitOnly(MethodInfo setMethod) =>
        setMethod.ReturnParameter
            .GetRequiredCustomModifiers()
            .Contains(typeof(IsExternalInit));

    static bool IsNullableReference(MemberInfo member, NullabilityInfoContext nullability)
    {
        var info = member switch
        {
            PropertyInfo property => nullability.Create(property),
            FieldInfo field => nullability.Create(field),
            _ => null
        };
        return info?.WriteState == NullabilityState.Nullable;
    }

    static EditableFieldKind? MapKind(Type type)
    {
        if (type == typeof(string))
        {
            return EditableFieldKind.Text;
        }

        if (type == typeof(bool))
        {
            return EditableFieldKind.Checkbox;
        }

        // DateOnly (aliased Date) and DateTime map to the native date picker, which stores a
        // canonical w:fullDate. DateTimeOffset and TimeOnly cannot: w:fullDate is a bare DateTime
        // with no offset, and Word has no time-only picker — so those render as round-trippable
        // plain text instead (see EditableFieldBuilder / EditableFieldReader).
        if (type == typeof(Date) ||
            type == typeof(DateTime))
        {
            return EditableFieldKind.Date;
        }

        if (type == typeof(DateTimeOffset))
        {
            return EditableFieldKind.DateTimeOffset;
        }

        if (type == typeof(Time))
        {
            return EditableFieldKind.Time;
        }

        if (type.IsEnum)
        {
            return EditableFieldKind.DropDown;
        }

        if (type == typeof(byte) ||
            type == typeof(sbyte) ||
            type == typeof(short) ||
            type == typeof(ushort) ||
            type == typeof(int) ||
            type == typeof(uint) ||
            type == typeof(long) ||
            type == typeof(ulong) ||
            type == typeof(float) ||
            type == typeof(double) ||
            type == typeof(decimal))
        {
            return EditableFieldKind.Number;
        }

        return null;
    }

    static Func<object, object?> ChainGetter(Func<object, object?> upstream, Func<object, object?> memberGetter) =>
        root =>
        {
            var parent = upstream(root);
            return parent == null ? null : memberGetter(parent);
        };

    static bool ShouldDescend(Type type)
    {
        if (type.IsPrimitive || type.IsEnum)
        {
            return false;
        }

        if (type == typeof(string) ||
            type == typeof(decimal) ||
            type == typeof(DateTime) ||
            type == typeof(DateTimeOffset) ||
            type == typeof(Date) ||
            type == typeof(Time) ||
            type == typeof(TimeSpan) ||
            type == typeof(Guid) ||
            type == typeof(Uri))
        {
            return false;
        }

        return !typeof(IEnumerable).IsAssignableFrom(type);
    }
}
