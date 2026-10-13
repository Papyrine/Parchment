/// <summary>
/// Builds a primitive-only <see cref="ModelShape"/> from a live <see cref="INamedTypeSymbol"/>
/// at extract time. Consuming the shape downstream (instead of the symbol) is what makes the
/// incremental pipeline actually cacheable.
/// The shape is rebuilt on every new compilation — <c>ForAttributeWithMetadataName</c> in
/// <see cref="ParchmentTemplateGenerator"/> re-runs the extract each time — so an edit to a model
/// type declared in a separate file is picked up; the rebuilt shape compares equal by value when
/// nothing a template can bind has changed, and downstream stages stay cached.
/// </summary>
static class ShapeBuilder
{
    public const string ExcelsiorTableAttributeFullName = "Parchment.ExcelsiorTableAttribute";
    public const string EditableFieldAttributeFullName = "Parchment.EditableFieldAttribute";

    static readonly SymbolDisplayFormat format = SymbolDisplayFormat.FullyQualifiedFormat;

    public static ModelShape Build(INamedTypeSymbol root, INamedTypeSymbol? excelsiorTableType, INamedTypeSymbol? editableFieldType, Cancel cancel)
    {
        var entries = ImmutableArray.CreateBuilder<TypeEntry>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var queue = new Queue<ITypeSymbol>();
        // Sorted, so the pragma the ids end up in reads the same whatever order the walk met them.
        var suppressed = new SortedSet<string>(StringComparer.Ordinal);
        var pending = new List<string>();

        Enqueue(root, visited, queue);
        while (queue.Count > 0)
        {
            cancel.ThrowIfCancellationRequested();
            var type = queue.Dequeue();
            entries.Add(BuildEntry(type, excelsiorTableType, editableFieldType, visited, queue, suppressed, pending));
        }

        return new(Fqn(root), new(entries.ToImmutable()), new(suppressed.ToImmutableArray()));
    }

    static TypeEntry BuildEntry(ITypeSymbol type, INamedTypeSymbol? excelsiorTableType, INamedTypeSymbol? editableFieldType, HashSet<string> visited, Queue<ITypeSymbol> queue, SortedSet<string> suppressed, List<string> pending)
    {
        // Every type in the shape is named by the generated code (a cast, a typeof), the root
        // included. One that is obsolete as an error only gets here as the root, where there is no
        // member to leave out.
        pending.Clear();
        UseSiteDiagnostics.TryCollectType(type, pending);
        suppressed.UnionWith(pending);

        string? elementFqn = null;
        if (type.SpecialType != SpecialType.System_String &&
            ModelSymbolResolver.TryGetElementType(type, out var element))
        {
            elementFqn = Fqn(element);
            Enqueue(element, visited, queue);
        }

        var liquid = LiquidKindFor(type);
        var members = ImmutableArray.CreateBuilder<MemberEntry>();

        // An enum's constants are public static fields, which is all a member needs to be here. But
        // a template never holds the enum: it is handed the text the value prints as, so a constant
        // recorded as a member validates `{{ Status.Final }}` and then renders nothing.
        if (liquid == LiquidKind.Enum)
        {
            return new(Fqn(type), elementFqn, new(members.ToImmutable()), Liquid: liquid);
        }

        // An array's CLR members (Length, IsFixedSize, Array.MaxLength…) are not binding surface,
        // and emitting accessors for them produces illegal casts like `string[].MaxLength`.
        if (type is IArrayTypeSymbol)
        {
            return new(Fqn(type), elementFqn, new(members.ToImmutable()), Liquid: liquid);
        }

        // KeyValuePair<K, V> is a System type but it's the iteration element for every
        // IDictionary<K, V> reached from the model. Surfacing Key and Value lets `{{ kv.Key }}`
        // validate and gives the accessor emission something to register, and enqueuing V (and K)
        // makes user-type values reachable.
        if (type is INamedTypeSymbol { IsGenericType: true } pair &&
            pair.OriginalDefinition.ToDisplayString() == "System.Collections.Generic.KeyValuePair<TKey, TValue>")
        {
            members.Add(new("Key", Fqn(pair.TypeArguments[0])));
            members.Add(new("Value", Fqn(pair.TypeArguments[1])));
            Enqueue(pair.TypeArguments[0], visited, queue);
            Enqueue(pair.TypeArguments[1], visited, queue);
            return new(Fqn(type), elementFqn, new(members.ToImmutable()), Liquid: liquid);
        }

        if (!IsSystemType(type))
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var current = type;
            while (current != null)
            {
                // A model's own collection class reaches a template as an array or a dictionary, so
                // none of its CLR members bind. Its own are still walked, since an attribute map
                // can reach them by getter, but what List<T> or Dictionary<K, V> declares beneath
                // it (Count, Capacity, Keys…) is not binding surface for anything.
                if (liquid != LiquidKind.None &&
                    IsSystemType(current))
                {
                    break;
                }

                foreach (var member in current.GetMembers())
                {
                    if (!TryGetMemberType(member, out var memberType, out var memberName, out var isStatic))
                    {
                        continue;
                    }

                    if (!seen.Add(memberName))
                    {
                        continue;
                    }

                    // A member obsolete as an error, or typed as something that is, has no legal
                    // reference: CS0619 cannot be suppressed. It is left out of the shape, so a
                    // token naming it is PARCH001 rather than a compile error in generated code.
                    // Anything milder stays bindable, with its diagnostic id carried to the pragma
                    // at the top of the generated file. The name is claimed first either way, so a
                    // base member it hides is not bound in its place.
                    pending.Clear();
                    if (!UseSiteDiagnostics.TryCollectMember(member, pending) ||
                        !UseSiteDiagnostics.TryCollectType(memberType, pending))
                    {
                        continue;
                    }

                    // An inherited member is read through its declaring type's attributes too: an
                    // experimental base class makes each of its members experimental.
                    if (!SymbolEqualityComparer.Default.Equals(current, type))
                    {
                        UseSiteDiagnostics.TryCollectType(current, pending);
                    }

                    suppressed.UnionWith(pending);

                    var isExcelsior = TryGetExcelsiorTable(member, excelsiorTableType, out var excelsiorHeadingStyle, out var excelsiorBodyStyle, out var excelsiorTableStyle, out var excelsiorConfigure, out var excelsiorConfigureOwner);
                    // Resolved here, where the declaring symbol is at hand: the generated code
                    // calls the method directly, so the name has to become a qualified target - and
                    // a name that resolves to no static method becomes the diagnostic instead. The
                    // type searched is the one the attribute was written in, which for an override
                    // inheriting it is the base that named the method.
                    string? excelsiorConfigureCall = null;
                    var excelsiorConfigureMissing = false;
                    if (excelsiorConfigure != null &&
                        excelsiorConfigureOwner != null)
                    {
                        var hasMethod = excelsiorConfigureOwner.GetMembers(excelsiorConfigure)
                            .OfType<IMethodSymbol>()
                            .Any(_ => _ is { IsStatic: true, Parameters.Length: 1 });
                        if (hasMethod)
                        {
                            excelsiorConfigureCall = $"{Fqn(excelsiorConfigureOwner)}.{excelsiorConfigure}";
                        }
                        else
                        {
                            // The raw name, for the diagnostic to quote.
                            excelsiorConfigureCall = excelsiorConfigure;
                            excelsiorConfigureMissing = true;
                        }
                    }
                    var (isHtml, isMarkdown, formatConflict) = DetectFormat(member);
                    // A TokenValue-typed member already declares its rendering by type, so a marker
                    // on top of it is the same claim made twice — and applying both converts the
                    // placeholder the first one produced. See PARCH025.
                    var markerOnToken = (isHtml || isMarkdown) && IsTokenValue(memberType);
                    // Decided here and carried on the entry, so the diagnostic (PARCH034) and
                    // the emission-side skip in AccessorEmission.WalkForMaps read one answer.
                    var excelsiorNotTabular = isExcelsior && !IsTabular(memberType);
                    // The conversion reads a string, so a marker on anything else has nothing to
                    // convert. An unresolved type is left to the compiler's own error. See PARCH029.
                    var markerOnNonString = (isHtml || isMarkdown) &&
                                            !markerOnToken &&
                                            memberType.SpecialType != SpecialType.System_String &&
                                            memberType.TypeKind != TypeKind.Error;
                    var isStringList = !isExcelsior &&
                                       IsEnumerableOfString(memberType);
                    var isEditable = TryGetEditableField(member, editableFieldType, out var editableMultiLine, out var editableDateFormat);

                    // An [EditableField] collection of a POCO element type renders as a repeating
                    // section: EditableKind stays null and the element fqn is recorded instead.
                    // A collection of a system element type (e.g. List<string>) is not one — it falls
                    // through to the scalar path, where MapEditableKind yields null → PARCH013.
                    string? editableCollectionElementFqn = null;
                    var editableCollectionAssignable = true;
                    EditableFieldKind? editableKind = null;
                    if (isEditable)
                    {
                        if (memberType.SpecialType != SpecialType.System_String &&
                            ModelSymbolResolver.TryGetElementType(memberType, out var collectionElement) &&
                            !IsSystemType(collectionElement))
                        {
                            editableCollectionElementFqn = Fqn(collectionElement);
                            editableCollectionAssignable = AcceptsRebuiltList(memberType);
                        }
                        else
                        {
                            editableKind = EditableKindFor(isHtml, isMarkdown, memberType);
                        }
                    }

                    members.Add(new(
                        memberName,
                        Fqn(memberType),
                        isExcelsior,
                        isHtml,
                        isMarkdown,
                        formatConflict,
                        markerOnToken,
                        isStringList,
                        isStatic,
                        excelsiorHeadingStyle,
                        excelsiorBodyStyle,
                        excelsiorTableStyle,
                        isEditable,
                        editableKind,
                        isEditable && editableCollectionElementFqn == null && IsNullableMember(memberType),
                        isEditable && HasUsableSetter(member),
                        editableMultiLine,
                        editableDateFormat,
                        editableCollectionElementFqn,
                        excelsiorConfigureCall,
                        excelsiorConfigureMissing,
                        excelsiorNotTabular,
                        markerOnNonString,
                        editableCollectionAssignable,
                        !isStatic && HasUsableSetter(member) && IsCreatable(memberType)));
                    Enqueue(memberType, visited, queue);
                }

                current = current.BaseType;
            }
        }

        return new(Fqn(type), elementFqn, new(members.ToImmutable()), HasParameterlessCtor(type), liquid);
    }

    // Lockstep with the runtime's ParchmentExtractor.ReconstructCollection, which rebuilds an
    // editable collection as a T[] for a member declared T[] and as a List<T> for everything else.
    // The generated setter casts that value to the declared type, so the declared type has to be
    // one of the two or an interface List<T> implements. Anything else (HashSet<T>, Collection<T>,
    // an immutable collection, a class deriving from List<T>) is PARCH022.
    static bool AcceptsRebuiltList(ITypeSymbol type)
    {
        if (type is IArrayTypeSymbol array)
        {
            return array.IsSZArray;
        }

        if (type is not INamedTypeSymbol {IsGenericType: true, TypeArguments.Length: 1} named)
        {
            return false;
        }

        return named.ConstructedFrom.ToDisplayString() is
            "System.Collections.Generic.List<T>" or
            "System.Collections.Generic.IList<T>" or
            "System.Collections.Generic.ICollection<T>" or
            "System.Collections.Generic.IEnumerable<T>" or
            "System.Collections.Generic.IReadOnlyList<T>" or
            "System.Collections.Generic.IReadOnlyCollection<T>";
    }

    // Lockstep with the runtime's LiquidMembers.Classify, which reads the same four cases off a
    // System.Type. The order is Fluid's: a dictionary is enumerable too, and is tested for first.
    static LiquidKind LiquidKindFor(ITypeSymbol type)
    {
        if (type.SpecialType == SpecialType.System_String)
        {
            return LiquidKind.String;
        }

        if (IsEnum(type))
        {
            return LiquidKind.Enum;
        }

        if (type is IArrayTypeSymbol)
        {
            return LiquidKind.Array;
        }

        var enumerable = type.SpecialType == SpecialType.System_Collections_IEnumerable;
        if (IsDictionary(type))
        {
            return LiquidKind.Dictionary;
        }

        foreach (var i in type.AllInterfaces)
        {
            if (IsDictionary(i))
            {
                return LiquidKind.Dictionary;
            }

            if (i.SpecialType == SpecialType.System_Collections_IEnumerable)
            {
                enumerable = true;
            }
        }

        if (enumerable)
        {
            return LiquidKind.Array;
        }

        return LiquidKind.None;
    }

    // A nullable enum is boxed on its way to Fluid, so what arrives is the enum or nothing.
    static bool IsEnum(ITypeSymbol type)
    {
        if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable)
        {
            type = nullable.TypeArguments[0];
        }

        return type.TypeKind == TypeKind.Enum;
    }

    static bool IsDictionary(ITypeSymbol type) =>
        type.OriginalDefinition.ToDisplayString() is
            "System.Collections.IDictionary" or
            "System.Collections.Generic.IDictionary<TKey, TValue>" or
            "System.Collections.Generic.IReadOnlyDictionary<TKey, TValue>";

    // Walked rather than compared to a referenced symbol: the shape is built from whatever the
    // consuming compilation has, and the base chain is enough to recognise one of the token types.
    static bool IsTokenValue(ITypeSymbol type)
    {
        for (var current = type; current != null; current = current.BaseType)
        {
            if (current.ToDisplayString(format) == "global::Parchment.TokenValue")
            {
                return true;
            }
        }

        return false;
    }

    static bool HasParameterlessCtor(ITypeSymbol type) =>
        type is INamedTypeSymbol named &&
        named.InstanceConstructors.Any(_ => _ is
        {
            Parameters.IsEmpty: true,
            DeclaredAccessibility: Accessibility.Public
        });

    static bool IsCreatable(ITypeSymbol type) =>
        type is INamedTypeSymbol { TypeKind: TypeKind.Class, IsAbstract: false } &&
        !IsSystemType(type) &&
        HasParameterlessCtor(type);

    static (bool isHtml, bool isMarkdown, bool conflict) DetectFormat(ISymbol member)
    {
        var hasHtml = false;
        var hasMarkdown = false;
        string? stringSyntax = null;
        foreach (var (attribute, _) in AttributesOf(member))
        {
            var cls = attribute.AttributeClass;
            if (cls == null)
            {
                continue;
            }

            // Both spellings, because the name is all there is to go on and it is not always the
            // declared one. When the attribute type is emitted by another source generator this one
            // cannot see it — generators all run against the same input compilation — so the symbol
            // arrives as an error type named exactly as written, "Html" rather than "HtmlAttribute".
            // Matching only the declared spelling made those markers vanish without a word, which is
            // how a whole report silently stopped rendering its html. Nothing is lost by accepting
            // the short form: this has only ever matched on the bare name, never on where the
            // attribute came from.
            // The base chain is walked so that an attribute deriving from one of the markers is
            // the marker too, as it is to anything reading attributes by type.
            if (IsOrDerivesFrom(cls, "HtmlAttribute", "Html"))
            {
                hasHtml = true;
            }
            else if (IsOrDerivesFrom(cls, "MarkdownAttribute", "Markdown"))
            {
                hasMarkdown = true;
            }
            else if (cls.ToDisplayString(format) == "global::System.Diagnostics.CodeAnalysis.StringSyntaxAttribute")
            {
                if (attribute.ConstructorArguments.Length > 0 &&
                    attribute.ConstructorArguments[0].Value is string value)
                {
                    stringSyntax = value.ToLowerInvariant();
                }
            }
        }

        // The markers must agree: [Html]+[Markdown], or either contradicted by [StringSyntax], is
        // PARCH023 — there is no principled winner to pick silently.
        var conflict = (hasHtml && hasMarkdown) ||
                       (hasHtml && stringSyntax == "markdown") ||
                       (hasMarkdown && stringSyntax == "html");

        if (hasHtml || stringSyntax == "html")
        {
            return (true, false, conflict);
        }

        if (hasMarkdown || stringSyntax == "markdown")
        {
            return (false, true, conflict);
        }

        return (false, false, false);
    }

    static bool IsOrDerivesFrom(INamedTypeSymbol cls, string name, string shortName)
    {
        for (var current = cls; current != null; current = current.BaseType)
        {
            if (current.Name == name ||
                current.Name == shortName)
            {
                return true;
            }
        }

        return false;
    }

    // What C# attribute inheritance gives a member: its own attributes, then those of each member
    // it overrides whose attribute class is inherited. The owner is the member the attribute was
    // written on, for what has to be resolved beside it. A member hiding another with `new`
    // overrides nothing, and an interface member's attributes are not inherited by its
    // implementation, so neither contributes.
    static IEnumerable<(AttributeData attribute, ISymbol owner)> AttributesOf(ISymbol member)
    {
        foreach (var attribute in member.GetAttributes())
        {
            yield return (attribute, member);
        }

        var overridden = (member as IPropertySymbol)?.OverriddenProperty;
        while (overridden != null)
        {
            foreach (var attribute in overridden.GetAttributes())
            {
                if (IsInherited(attribute.AttributeClass))
                {
                    yield return (attribute, overridden);
                }
            }

            overridden = overridden.OverriddenProperty;
        }
    }

    // [AttributeUsage] is itself inherited, so the nearest one up the attribute's base chain
    // decides, and Inherited is true unless it says otherwise.
    static bool IsInherited(INamedTypeSymbol? cls)
    {
        for (var current = cls; current != null; current = current.BaseType)
        {
            foreach (var attribute in current.GetAttributes())
            {
                if (attribute.AttributeClass?.ToDisplayString(format) != "global::System.AttributeUsageAttribute")
                {
                    continue;
                }

                foreach (var named in attribute.NamedArguments)
                {
                    if (named is { Key: "Inherited", Value.Value: bool inherited })
                    {
                        return inherited;
                    }
                }

                return true;
            }
        }

        return true;
    }

    // `string` itself is `IEnumerable<char>`, not `IEnumerable<string>` — element type would
    // be `char`, which is correctly rejected here.
    static bool IsEnumerableOfString(ITypeSymbol type) =>
        TryGetSequenceElementType(type, out var element) &&
        element is {SpecialType: SpecialType.System_String};

    // A multi-dimensional array has an element type but is no IEnumerable<T>: the runtime reads a
    // string list and an Excelsior source through that interface, so `string[,]` handed to either
    // yields nothing. A jagged `string[][]` is a rank-1 array of arrays and is unaffected.
    static bool TryGetSequenceElementType(ITypeSymbol type, [NotNullWhen(true)] out ITypeSymbol? element)
    {
        element = null;
        if (type.SpecialType == SpecialType.System_String ||
            type is IArrayTypeSymbol {Rank: > 1})
        {
            return false;
        }

        return ModelSymbolResolver.TryGetElementType(type, out element);
    }

    /// <summary>
    /// Whether an <c>[ExcelsiorTable]</c> member's type is a sequence of something Excelsior can
    /// lay out as columns: an element with at least one public instance property or field, which
    /// is what its column discovery reads. A scalar, a string, an untyped <c>IEnumerable</c>, and a
    /// sequence of strings, primitives, enums, <c>object</c> or <c>dynamic</c> have none.
    /// </summary>
    static bool IsTabular(ITypeSymbol type)
    {
        if (!TryGetSequenceElementType(type, out var element))
        {
            return false;
        }

        // Covers string, object, the primitives, decimal, DateTime and Nullable<T>.
        if (element.OriginalDefinition.SpecialType != SpecialType.None)
        {
            return false;
        }

        for (var current = element; current != null; current = current.BaseType)
        {
            foreach (var member in current.GetMembers())
            {
                if (TryGetMemberType(member, out _, out _, out var isStatic) &&
                    !isStatic)
                {
                    return true;
                }
            }
        }

        return false;
    }

    static bool TryGetExcelsiorTable(ISymbol member, INamedTypeSymbol? excelsiorTableType, out string? headingParagraphStyle, out string? bodyParagraphStyle, out string? tableStyle, out string? configure, out INamedTypeSymbol? configureOwner)
    {
        headingParagraphStyle = null;
        bodyParagraphStyle = null;
        tableStyle = null;
        configure = null;
        configureOwner = null;
        if (excelsiorTableType is null)
        {
            return false;
        }

        foreach (var (attribute, owner) in AttributesOf(member))
        {
            if (!SymbolEqualityComparer.Default.Equals(attribute.AttributeClass, excelsiorTableType))
            {
                continue;
            }

            foreach (var named in attribute.NamedArguments)
            {
                if (named.Value.Value is not string value)
                {
                    continue;
                }

                // String literals (not nameof) — the SG can't reference Parchment.dll, so it has no
                // typed handle on ExcelsiorTableAttribute (matched by FQN string elsewhere too).
                if (named.Key == "HeadingParagraphStyle")
                {
                    headingParagraphStyle = value;
                }
                else if (named.Key == "BodyParagraphStyle")
                {
                    bodyParagraphStyle = value;
                }
                else if (named.Key == "TableStyle")
                {
                    tableStyle = value;
                }
                else if (named.Key == "Configure")
                {
                    configure = value;
                    configureOwner = owner.ContainingType;
                }
            }

            return true;
        }

        return false;
    }

    static bool TryGetEditableField(ISymbol member, INamedTypeSymbol? editableFieldType, out bool multiLine, out string? dateFormat)
    {
        multiLine = false;
        dateFormat = null;
        if (editableFieldType is null)
        {
            return false;
        }

        foreach (var (attribute, _) in AttributesOf(member))
        {
            if (!SymbolEqualityComparer.Default.Equals(attribute.AttributeClass, editableFieldType))
            {
                continue;
            }

            foreach (var named in attribute.NamedArguments)
            {
                if (named is { Key: "MultiLine", Value.Value: bool multi })
                {
                    multiLine = multi;
                }
                else if (named is { Key: "DateFormat", Value.Value: string dateFormatValue })
                {
                    dateFormat = dateFormatValue;
                }
            }

            return true;
        }

        return false;
    }

    /// <summary>
    /// Chooses the editable kind for a member, accounting for the format markers. Lockstep with
    /// runtime <c>EditableMap.BuildEntry</c>: <c>[Html]</c> on a string selects the rich-content
    /// <see cref="EditableFieldKind.Html"/>; <c>[Markdown]</c> yields null so it's never emitted as
    /// an editable member (the PARCH015 conflict pass reports it first).
    /// </summary>
    static EditableFieldKind? EditableKindFor(bool isHtml, bool isMarkdown, ITypeSymbol type)
    {
        if (isMarkdown)
        {
            return null;
        }

        if (isHtml)
        {
            if (type.SpecialType == SpecialType.System_String)
            {
                return EditableFieldKind.Html;
            }

            return null;
        }

        return MapEditableKind(type);
    }

    /// <summary>
    /// The runtime lockstep is <c>EditableMap.MapKind</c> + the bool? guard: null means
    /// PARCH013 (unsupported type — including <c>bool?</c>, which a checkbox cannot represent, and
    /// a <c>[Flags]</c> enum, which a dropdown cannot).
    /// </summary>
    static EditableFieldKind? MapEditableKind(ITypeSymbol type)
    {
        var isNullableValue = false;
        if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullable)
        {
            isNullableValue = true;
            type = nullable.TypeArguments[0];
        }

        switch (type.SpecialType)
        {
            case SpecialType.System_String:
                return EditableFieldKind.Text;
            case SpecialType.System_Boolean:
                if (isNullableValue)
                {
                    return null;
                }

                return EditableFieldKind.Checkbox;
            case SpecialType.System_DateTime:
                return EditableFieldKind.Date;
            case SpecialType.System_Byte:
            case SpecialType.System_SByte:
            case SpecialType.System_Int16:
            case SpecialType.System_UInt16:
            case SpecialType.System_Int32:
            case SpecialType.System_UInt32:
            case SpecialType.System_Int64:
            case SpecialType.System_UInt64:
            case SpecialType.System_Single:
            case SpecialType.System_Double:
            case SpecialType.System_Decimal:
                return EditableFieldKind.Number;
        }

        if (type.TypeKind == TypeKind.Enum)
        {
            // A dropdown selects one item, and a [Flags] value is a combination: Read | Write
            // matches no item, so the control would show text that is not in its list and
            // extraction could not read it back. The attribute decides, not the constants.
            if (IsFlagsEnum(type))
            {
                return null;
            }

            return EditableFieldKind.DropDown;
        }

        // DateOnly + DateTime -> native date picker (canonical w:fullDate). DateTimeOffset and
        // TimeOnly -> round-trippable plain text (w:fullDate has no offset; no time-only picker).
        // Lockstep with runtime EditableMap.MapKind.
        var fqn = Fqn(type);
        return fqn switch
        {
            "global::System.DateOnly" => EditableFieldKind.Date,
            "global::System.DateTimeOffset" => EditableFieldKind.DateTimeOffset,
            "global::System.TimeOnly" => EditableFieldKind.Time,
            _ => null
        };
    }

    static bool IsFlagsEnum(ITypeSymbol type)
    {
        foreach (var attribute in type.GetAttributes())
        {
            if (attribute.AttributeClass is { Name: "FlagsAttribute", ContainingNamespace: { Name: "System", ContainingNamespace.IsGlobalNamespace: true } })
            {
                return true;
            }
        }

        return false;
    }

    static bool IsNullableMember(ITypeSymbol type) =>
        type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } ||
        type.NullableAnnotation == NullableAnnotation.Annotated;

    static bool HasUsableSetter(ISymbol member) =>
        member switch
        {
            IPropertySymbol { SetMethod: { IsInitOnly: false, DeclaredAccessibility: Accessibility.Public } } => true,
            IFieldSymbol { IsReadOnly: false, IsConst: false } => true,
            _ => false
        };

    static bool TryGetMemberType(ISymbol member, out ITypeSymbol type, out string name, out bool isStatic)
    {
        // An indexer has no name a template can ask for, and a property without a public getter has
        // nothing the generated code can read: an accessor for either (`o.this[]`, or a call to a
        // private get) does not compile. Lockstep with the runtime's ModelValidator.TryResolveMember.
        if (member is IPropertySymbol { DeclaredAccessibility: Accessibility.Public, IsIndexer: false, GetMethod.DeclaredAccessibility: Accessibility.Public } property)
        {
            type = property.Type;
            name = property.Name;
            isStatic = property.IsStatic;
            return true;
        }

        if (member is IFieldSymbol { DeclaredAccessibility: Accessibility.Public } field)
        {
            type = field.Type;
            name = field.Name;
            isStatic = field.IsStatic;
            return true;
        }

        type = null!;
        name = null!;
        isStatic = false;
        return false;
    }

    static void Enqueue(ITypeSymbol? type, HashSet<string> visited, Queue<ITypeSymbol> queue)
    {
        if (type == null)
        {
            return;
        }

        var key = Fqn(type);
        if (visited.Add(key))
        {
            queue.Enqueue(type);
        }
    }

    static bool IsSystemType(ITypeSymbol type)
    {
        var ns = type.ContainingNamespace;
        if (ns is null or { IsGlobalNamespace: true })
        {
            return false;
        }

        while (ns.ContainingNamespace is { IsGlobalNamespace: false })
        {
            ns = ns.ContainingNamespace;
        }

        return ns.Name == "System";
    }

    static string Fqn(ITypeSymbol type) =>
        type.ToDisplayString(format);
}
