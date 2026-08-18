/// <summary>
/// Builds a primitive-only <see cref="ModelShape"/> from a live <see cref="INamedTypeSymbol"/>
/// at extract time. Consuming the shape downstream (instead of the symbol) is what makes the
/// incremental pipeline actually cacheable.
/// Known limitation: the shape is only rebuilt when the attributed class's own syntax changes,
/// because <c>ForAttributeWithMetadataName</c> in <see cref="ParchmentTemplateGenerator"/> keys
/// re-extraction on that class's syntax. Edits to a model type declared in a separate file
/// will not re-trigger validation until something in the attributed class is touched.
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

        Enqueue(root, visited, queue);
        while (queue.Count > 0)
        {
            cancel.ThrowIfCancellationRequested();
            var type = queue.Dequeue();
            entries.Add(BuildEntry(type, excelsiorTableType, editableFieldType, visited, queue));
        }

        return new(Fqn(root), new(entries.ToImmutable()));
    }

    static TypeEntry BuildEntry(ITypeSymbol type, INamedTypeSymbol? excelsiorTableType, INamedTypeSymbol? editableFieldType, HashSet<string> visited, Queue<ITypeSymbol> queue)
    {
        string? elementFqn = null;
        if (type.SpecialType != SpecialType.System_String &&
            ModelSymbolResolver.TryGetElementType(type, out var element))
        {
            elementFqn = Fqn(element);
            Enqueue(element, visited, queue);
        }

        var members = ImmutableArray.CreateBuilder<MemberEntry>();

        // An array's CLR members (Length, IsFixedSize, Array.MaxLength…) are not binding surface,
        // and emitting accessors for them produces illegal casts like `string[].MaxLength`.
        if (type is IArrayTypeSymbol)
        {
            return new(Fqn(type), elementFqn, new(members.ToImmutable()));
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
            return new(Fqn(type), elementFqn, new(members.ToImmutable()));
        }

        if (!IsSystemType(type))
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var current = type;
            while (current != null)
            {
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

                    var isExcelsior = TryGetExcelsiorTable(member, excelsiorTableType, out var excelsiorHeadingStyle, out var excelsiorBodyStyle, out var excelsiorTableStyle, out var excelsiorConfigure);
                    // Resolved here, where the declaring symbol is at hand: the generated code
                    // calls the method directly, so the name has to become a qualified target - and
                    // a name that resolves to no static method becomes the diagnostic instead.
                    string? excelsiorConfigureCall = null;
                    var excelsiorConfigureMissing = false;
                    if (excelsiorConfigure != null)
                    {
                        var hasMethod = current.GetMembers(excelsiorConfigure)
                            .OfType<IMethodSymbol>()
                            .Any(_ => _.IsStatic && _.Parameters.Length == 1);
                        if (hasMethod)
                        {
                            excelsiorConfigureCall = $"{Fqn(current)}.{excelsiorConfigure}";
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
                    var isStringList = !isExcelsior &&
                                       IsEnumerableOfString(memberType);
                    var isEditable = TryGetEditableField(member, editableFieldType, out var editableMultiLine, out var editableDateFormat);

                    // An [EditableField] collection of a POCO element type renders as a repeating
                    // section: EditableKind stays null and the element fqn is recorded instead.
                    // A collection of a system element type (e.g. List<string>) is not one — it falls
                    // through to the scalar path, where MapEditableKind yields null → PARCH013.
                    string? editableCollectionElementFqn = null;
                    EditableFieldKind? editableKind = null;
                    if (isEditable)
                    {
                        if (memberType.SpecialType != SpecialType.System_String &&
                            ModelSymbolResolver.TryGetElementType(memberType, out var collectionElement) &&
                            !IsSystemType(collectionElement))
                        {
                            editableCollectionElementFqn = Fqn(collectionElement);
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
                        excelsiorConfigureMissing));
                    Enqueue(memberType, visited, queue);
                }

                current = current.BaseType;
            }
        }

        return new(Fqn(type), elementFqn, new(members.ToImmutable()), HasParameterlessCtor(type));
    }

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

    static (bool isHtml, bool isMarkdown, bool conflict) DetectFormat(ISymbol member)
    {
        var hasHtml = false;
        var hasMarkdown = false;
        string? stringSyntax = null;
        foreach (var attribute in member.GetAttributes())
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
            var name = cls.Name;
            if (name is "HtmlAttribute" or "Html")
            {
                hasHtml = true;
            }
            else if (name is "MarkdownAttribute" or "Markdown")
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

    // `string` itself is `IEnumerable<char>`, not `IEnumerable<string>` — element type would
    // be `char`, which is correctly rejected here.
    static bool IsEnumerableOfString(ITypeSymbol type) =>
        ModelSymbolResolver.TryGetElementType(type, out var element) &&
        element is {SpecialType: SpecialType.System_String};

    static bool TryGetExcelsiorTable(ISymbol member, INamedTypeSymbol? excelsiorTableType, out string? headingParagraphStyle, out string? bodyParagraphStyle, out string? tableStyle, out string? configure)
    {
        headingParagraphStyle = null;
        bodyParagraphStyle = null;
        tableStyle = null;
        configure = null;
        if (excelsiorTableType is null)
        {
            return false;
        }

        foreach (var attribute in member.GetAttributes())
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

        foreach (var attribute in member.GetAttributes())
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
            return type.SpecialType == SpecialType.System_String ? EditableFieldKind.Html : null;
        }

        return MapEditableKind(type);
    }

    /// <summary>
    /// The runtime lockstep is <c>EditableMap.MapKind</c> + the bool? guard: null means
    /// PARCH013 (unsupported type — including <c>bool?</c>, which a checkbox cannot represent).
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
                return isNullableValue ? null : EditableFieldKind.Checkbox;
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
        if (member is IPropertySymbol { DeclaredAccessibility: Accessibility.Public } property)
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
