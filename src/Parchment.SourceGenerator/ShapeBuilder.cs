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
        if (type.SpecialType != SpecialType.System_String)
        {
            var element = ModelSymbolResolver.TryGetElementType(type);
            if (element != null)
            {
                elementFqn = Fqn(element);
                Enqueue(element, visited, queue);
            }
        }

        var members = ImmutableArray.CreateBuilder<MemberEntry>();
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

                    var isExcelsior = TryGetExcelsiorTable(member, excelsiorTableType, out var excelsiorHeadingStyle, out var excelsiorBodyStyle);
                    var (isHtml, isMarkdown) = DetectFormat(member);
                    var isStringList = !isExcelsior &&
                                       IsEnumerableOfString(memberType);
                    var isEditable = TryGetEditableField(member, editableFieldType, out var editableMultiLine, out var editableDateFormat);
                    members.Add(new(
                        memberName,
                        Fqn(memberType),
                        isExcelsior,
                        isHtml,
                        isMarkdown,
                        isStringList,
                        isStatic,
                        excelsiorHeadingStyle,
                        excelsiorBodyStyle,
                        isEditable,
                        isEditable ? MapEditableKind(memberType) : null,
                        isEditable && IsNullableMember(memberType),
                        isEditable && HasUsableSetter(member),
                        editableMultiLine,
                        editableDateFormat));
                    Enqueue(memberType, visited, queue);
                }

                current = current.BaseType;
            }
        }

        return new(Fqn(type), elementFqn, new(members.ToImmutable()));
    }

    static (bool isHtml, bool isMarkdown) DetectFormat(ISymbol member)
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

            var name = cls.Name;
            if (name == "HtmlAttribute")
            {
                hasHtml = true;
            }
            else if (name == "MarkdownAttribute")
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

        if (hasHtml || stringSyntax == "html")
        {
            return (true, false);
        }

        if (hasMarkdown || stringSyntax == "markdown")
        {
            return (false, true);
        }

        return (false, false);
    }

    static bool IsEnumerableOfString(ITypeSymbol type)
    {
        // `string` itself is `IEnumerable<char>`, not `IEnumerable<string>` — element type would
        // be `char`, which is correctly rejected here.
        var element = ModelSymbolResolver.TryGetElementType(type);
        return element is {SpecialType: SpecialType.System_String};
    }

    static bool TryGetExcelsiorTable(ISymbol member, INamedTypeSymbol? excelsiorTableType, out string? headingParagraphStyle, out string? bodyParagraphStyle)
    {
        headingParagraphStyle = null;
        bodyParagraphStyle = null;
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
