/// <summary>
/// Source-generator emission of pre-compiled registration data covering every reachable type
/// in the model graph. Produces four parallel datasets so the runtime's reflection walks
/// (<c>SharedFluid.RegisterTypeGraph</c>, <c>ExcelsiorTableMap.WalkType</c>,
/// <c>FormatMap.WalkType</c>, <c>StringListMap.WalkType</c>) can be skipped entirely on the
/// SG path:
///
/// <list type="bullet">
///   <item><description>Fluid accessors per registered type — one <c>DelegateAccessor</c> per public
///     property / field, so Fluid's <c>MemberAccessStrategy</c> can look up tokens without
///     reflection.</description></item>
///   <item><description>Excelsior-table entries — dotted path + element type + getter for every
///     <c>[ExcelsiorTable]</c>-marked property reachable from the root model.</description></item>
///   <item><description>Format entries — dotted path + <c>FormatMapKind</c> + getter for every
///     <c>[Html]</c> / <c>[Markdown]</c> / <c>[StringSyntax]</c>-marked string property.</description></item>
///   <item><description>String-list entries — dotted path + getter for every
///     <c>IEnumerable&lt;string&gt;</c> property that isn't already owned by
///     <c>[ExcelsiorTable]</c>.</description></item>
///   <item><description>Editable entries — dotted path + kind + getter + setter + reachability
///     check for every valid <c>[EditableField]</c> member. Setters route through the parent
///     object (never the root — the leaf lives on the object at the end of the path), matching
///     the runtime <c>EditableMap.BuildSetter</c>.</description></item>
/// </list>
///
/// Path-getters are emitted as inline cast-and-null-chain lambdas:
/// <c>o => ((global::NS.Root)o).First?.Second?.Third</c>. Works for reference-type chains only.
/// Value-type intermediates are unsupported (rare in real binding models — value types are
/// typically leaves): <c>?.</c> does not compile against a non-nullable struct, and a
/// <c>Nullable&lt;T&gt;</c> member is never walked into — <c>ShapeBuilder</c> gives a System type
/// no members, so nothing beneath it resolves or reaches a map.
/// </summary>

static class AccessorEmission
{
    public sealed record EmittedAccessors(
        string FieldsBlock,
        string RegistrationsBlock);

    public static EmittedAccessors? Emit(ModelShape shape, string rootFqn)
    {
        var typesByFqn = new Dictionary<string, TypeEntry>(StringComparer.Ordinal);
        foreach (var type in shape.Types)
        {
            typesByFqn[type.TypeFullyQualifiedName] = type;
        }

        var fluidBlocks = new List<(string FieldName, TypeEntry Type)>();
        var excelsiorEntries = new List<(List<string> Path, string ElementFqn, string? HeadingParagraphStyle, string? BodyParagraphStyle, string? TableStyle, string? ConfigureCall)>();
        var formatEntries = new List<(List<string> Path, FormatMapKind Kind)>();
        var stringListEntries = new List<List<string>>();
        var editableEntries = new List<(List<string> Path, MemberEntry Member)>();
        var collectionEntries = new List<(List<string> Path, MemberEntry Member)>();

        // Fluid accessors: one block per non-empty type in the shape (System types & enums end
        // up with 0 members and are skipped). The root type is always emitted, even member-less —
        // its registration is what marks the model SG-known to SharedFluid.EnsureModelRegistered.
        foreach (var type in shape.Types)
        {
            if (type.Members.Count == 0 &&
                type.TypeFullyQualifiedName != rootFqn)
            {
                continue;
            }

            fluidBlocks.Add(("_Accessors_" + ToIdentifier(type.TypeFullyQualifiedName), type));
        }

        // Excelsior / Format / StringList walks: dotted-path traversal from the root TModel,
        // collecting attribute-marked properties. Per-branch visited tracking mirrors the runtime
        // walkers, so self-referential models terminate.
        WalkForMaps(
            rootFqn,
            [],
            new(StringComparer.Ordinal)
            {
                rootFqn
            },
            typesByFqn,
            excelsiorEntries,
            formatEntries,
            stringListEntries,
            editableEntries,
            collectionEntries,
            [],
            null);

        if (fluidBlocks.Count == 0 &&
            excelsiorEntries.Count == 0 &&
            formatEntries.Count == 0 &&
            stringListEntries.Count == 0 &&
            editableEntries.Count == 0 &&
            collectionEntries.Count == 0)
        {
            return null;
        }

        var fields = new StringBuilder();
        var registrations = new StringBuilder();

        EmitFluidBlocks(fields, registrations, fluidBlocks);
        EmitExcelsiorBlock(fields, registrations, rootFqn, excelsiorEntries);
        EmitFormatBlock(fields, registrations, rootFqn, formatEntries);
        EmitStringListBlock(fields, registrations, rootFqn, stringListEntries);
        EmitEditableBlock(fields, registrations, rootFqn, editableEntries);
        EmitCollectionBlock(fields, registrations, rootFqn, typesByFqn, collectionEntries);

        fields.TrimTrailingNewlines();
        registrations.TrimTrailingNewlines();

        return new(fields.ToString(), registrations.ToString());
    }

    // Null for a name that resolved to nothing callable: the diagnostic is then the only error.
    static string? ConfigureCall(MemberEntry member)
    {
        if (member.ExcelsiorConfigureMissing ||
            member.ExcelsiorConfigureInaccessibleOn != null)
        {
            return null;
        }

        return member.ExcelsiorConfigure;
    }

    /// <summary>
    /// The partial parts that give the registration a way to call each Configure method it could
    /// not reach, or null when every one is called directly.
    /// </summary>
    /// <remarks>
    /// A file of its own, since each part goes in its declaring type's namespace and the model's
    /// file is committed to the model's by a file-scoped declaration. Nullable is left unset, as
    /// the parameter is restated without its annotations. The forwarder is named for the model so
    /// two models reaching one type do not both declare the same member.
    /// </remarks>
    public static string? EmitConfigureForwarders(ModelShape shape)
    {
        var seen = new HashSet<ConfigureForwarder>();
        var builder = new StringBuilder();
        foreach (var type in shape.Types)
        {
            foreach (var member in type.Members)
            {
                foreach (var forwarder in member.ExcelsiorConfigureForwarders)
                {
                    if (seen.Add(forwarder))
                    {
                        EmitConfigureForwarder(builder, forwarder);
                    }
                }
            }
        }

        if (seen.Count == 0)
        {
            return null;
        }

        return "// <auto-generated />\n" + builder;
    }

    static void EmitConfigureForwarder(StringBuilder builder, ConfigureForwarder forwarder)
    {
        builder.Append('\n');
        var depth = 0;
        if (forwarder.Namespace != null)
        {
            builder.Append($"namespace {forwarder.Namespace}\n{{\n");
            depth++;
        }

        foreach (var type in forwarder.Types)
        {
            builder.Indent(depth).Append($"partial {type.Kind} {type.Name}\n");
            builder.Indent(depth).Append("{\n");
            depth++;
        }

        builder.Indent(depth).Append("[global::System.ComponentModel.EditorBrowsable(global::System.ComponentModel.EditorBrowsableState.Never)]\n");
        builder.Indent(depth).Append($"internal static void {forwarder.Name}({forwarder.ParameterType} builder) =>\n");
        builder.Indent(depth + 1).Append($"{forwarder.Method}(builder);\n");

        while (depth > 0)
        {
            depth--;
            builder.Indent(depth).Append("}\n");
        }
    }

    static void WalkForMaps(
        string currentTypeFqn,
        List<string> path,
        HashSet<string> visited,
        Dictionary<string, TypeEntry> typesByFqn,
        List<(List<string>, string, string?, string?, string?, string?)> excelsior,
        List<(List<string>, FormatMapKind)> formats,
        List<List<string>> stringLists,
        List<(List<string>, MemberEntry)> editables,
        List<(List<string>, MemberEntry)> collections,
        List<(string DeclaringTypeFqn, MemberEntry Member, string ValueTypeFqn)> throughValueType,
        string? valueTypeFqn)
    {
        if (!typesByFqn.TryGetValue(currentTypeFqn, out var typeEntry))
        {
            return;
        }

        // The first struct on the way here, the current type included. An editable member below it
        // has no setter that reaches the model: assigning through a struct hop assigns to a copy
        // (and through a boxed struct does not compile), so it is reported (PARCH032) rather than
        // emitted.
        if (typeEntry.IsValueType)
        {
            valueTypeFqn ??= typeEntry.TypeFullyQualifiedName;
        }

        foreach (var member in typeEntry.Members)
        {
            // The per-template maps (Excelsior / Format / StringList) match the runtime walkers,
            // which use BindingFlags.Instance only. Static members are still bindable through
            // Fluid accessors (emitted above), but they don't participate in dotted-path map
            // dispatch — `[ExcelsiorTable]` on a static collection, `[Html]` on a static string,
            // etc. is intentionally a no-op. Document if a user request comes up.
            if (member.IsStatic)
            {
                continue;
            }

            var nextPath = new List<string>(path)
            {
                member.Name
            };

            if (member.IsExcelsiorTable)
            {
                if (!member.ExcelsiorNotTabular &&
                    typesByFqn.TryGetValue(member.TypeFullyQualifiedName, out var memberType) &&
                    memberType.ElementTypeFullyQualifiedName != null)
                {
                    excelsior.Add((nextPath, memberType.ElementTypeFullyQualifiedName, member.ExcelsiorHeadingParagraphStyle, member.ExcelsiorBodyParagraphStyle, member.ExcelsiorTableStyle, ConfigureCall(member)));
                }

                continue;
            }

            if (member.IsStringList)
            {
                stringLists.Add(nextPath);
                continue;
            }

            // A member that is also [EditableField] is claimed by the editable map (rich-content
            // control that extracts back to HTML), not the read-only format path. A conflicted
            // member (PARCH023) is emitted as neither — the diagnostic fails the build. Nor is a
            // marker on a member that is not a string (PARCH029).
            if (member is {IsHtml: true, IsEditable: false, HasFormatConflict: false, FormatMarkerOnTokenValue: false, FormatMarkerOnNonString: false})
            {
                formats.Add((nextPath, FormatMapKind.Html));
                continue;
            }

            if (member is {IsMarkdown: true, IsEditable: false, HasFormatConflict: false, FormatMarkerOnTokenValue: false, FormatMarkerOnNonString: false})
            {
                formats.Add((nextPath, FormatMapKind.Markdown));
                continue;
            }

            if (member.IsEditable)
            {
                if (valueTypeFqn != null)
                {
                    throughValueType.Add((typeEntry.TypeFullyQualifiedName, member, valueTypeFqn));
                    continue;
                }

                // An editable collection (POCO element) → repeating section, registered separately.
                // Invalid shapes (no setter, a declared type the rebuilt list cannot be assigned
                // to, a struct element, element without a parameterless ctor or without editable members) carry an error diagnostic (PARCH014 / PARCH022) — skip them so
                // the emitted source stays compilable.
                if (member.EditableCollectionElementFqn != null)
                {
                    if (member.HasUsableSetter &&
                        member.EditableCollectionAssignable &&
                        typesByFqn.TryGetValue(member.EditableCollectionElementFqn, out var elementEntry) &&
                        elementEntry is { HasParameterlessCtor: true, IsValueType: false } &&
                        HasScalarEditableMember(elementEntry))
                    {
                        collections.Add((nextPath, member));
                    }

                    continue;
                }

                // Invalid editable members (unsupported type, no setter) already carry an error
                // diagnostic (PARCH013/014) — skip them so the emitted source stays compilable.
                if (member is { EditableKind: not null, HasUsableSetter: true })
                {
                    editables.Add((nextPath, member));
                }

                continue;
            }

            // Descend only into POCO branches — types with members. Enumerables and system types
            // end up with 0 members in the shape, so this filter naturally skips them.
            if (!typesByFqn.TryGetValue(member.TypeFullyQualifiedName, out var nextType) ||
                nextType.Members.Count == 0)
            {
                continue;
            }

            if (!visited.Add(member.TypeFullyQualifiedName))
            {
                continue;
            }

            WalkForMaps(member.TypeFullyQualifiedName, nextPath, visited, typesByFqn, excelsior, formats, stringLists, editables, collections, throughValueType, valueTypeFqn);
            visited.Remove(member.TypeFullyQualifiedName);
        }
    }

    /// <summary>
    /// The editable members the walks above leave out because a struct sits on the path to them,
    /// from the root and from each editable collection's element. The same walk that emits, so
    /// what is reported (PARCH032) is exactly what is not emitted.
    /// </summary>
    public static List<(string DeclaringTypeFqn, MemberEntry Member, string ValueTypeFqn)> EditablesThroughValueType(ModelShape shape)
    {
        var typesByFqn = new Dictionary<string, TypeEntry>(StringComparer.Ordinal);
        foreach (var type in shape.Types)
        {
            typesByFqn[type.TypeFullyQualifiedName] = type;
        }

        var result = new List<(string, MemberEntry, string)>();
        var collections = new List<(List<string> Path, MemberEntry Member)>();
        Walk(shape.RootTypeFullyQualifiedName, collections);
        foreach (var (_, member) in collections)
        {
            Walk(member.EditableCollectionElementFqn!, []);
        }

        return result;

        void Walk(string fqn, List<(List<string>, MemberEntry)> found) =>
            WalkForMaps(
                fqn,
                [],
                new(StringComparer.Ordinal)
                {
                    fqn
                },
                typesByFqn,
                [],
                [],
                [],
                [],
                found,
                result,
                null);
    }

    static bool HasScalarEditableMember(TypeEntry type)
    {
        foreach (var member in type.Members)
        {
            if (member is { IsEditable: true, IsStatic: false, EditableCollectionElementFqn: null })
            {
                return true;
            }
        }

        return false;
    }

    static void EmitFluidBlocks(
        StringBuilder fields,
        StringBuilder registrations,
        List<(string FieldName, TypeEntry Type)> blocks)
    {
        foreach (var (fieldName, type) in blocks)
        {
            fields.Append(
                $$"""
                static readonly global::System.Collections.Generic.KeyValuePair<string, global::Fluid.IMemberAccessor>[] {{fieldName}} =
                {

                """);
            foreach (var member in type.Members)
            {
                fields.Append(
                    $"""  new("{member.Name}", new global::Fluid.Accessors.DelegateAccessor(""");
                if (member.IsStatic)
                {
                    // Static member: the instance argument is irrelevant. Access via the owner
                    // type directly so the emitted code is legal C# (CS0176 forbids instance-
                    // qualifying a static member).
                    fields.Append($"(_, _) => {type.TypeFullyQualifiedName}.{member.Name}");
                }
                else
                {
                    fields.Append($"(o, _) => (({type.TypeFullyQualifiedName})o).{member.Name}");
                }

                fields.AppendLine(")),");
            }

            fields.Append(
                """
                };

                """);

            registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterFluidAccessors(typeof({type.TypeFullyQualifiedName}), {fieldName});");
        }
    }

    static void EmitExcelsiorBlock(
        StringBuilder fields,
        StringBuilder registrations,
        string rootFqn,
        List<(List<string> Path, string ElementFqn, string? HeadingParagraphStyle, string? BodyParagraphStyle, string? TableStyle, string? ConfigureCall)> entries)
    {
        if (entries.Count == 0)
        {
            return;
        }

        fields.Append(
            """
            static readonly global::Parchment.Generated.ExcelsiorTableMapEntry[] _ExcelsiorTables =
            {

            """);
        foreach (var (path, elementFqn, headingParagraphStyle, bodyParagraphStyle, tableStyle, configureCall) in entries)
        {
            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", typeof(");
            fields.Append(elementFqn);
            fields.Append("), ");
            EmitGetter(fields, rootFqn, path);
            fields.Append(", ");
            fields.Append(ToStringLiteral(headingParagraphStyle));
            fields.Append(", ");
            fields.Append(ToStringLiteral(bodyParagraphStyle));
            fields.Append(", ");
            fields.Append(ToStringLiteral(tableStyle));
            if (configureCall != null)
            {
                // A direct call rather than a name resolved at render time: a wrong signature is an
                // ordinary compile error here, and the method itself stays reachable by nameof.
                fields.Append(", b => ");
                fields.Append(configureCall);
                fields.Append("((global::Excelsior.WordTableBuilder<");
                fields.Append(elementFqn);
                fields.Append(">)b)");
            }

            fields.AppendLine("),");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterExcelsiorTable(typeof({rootFqn}), _ExcelsiorTables);");
    }

    static string ToStringLiteral(string? value)
    {
        if (value == null)
        {
            return "null";
        }

        return SymbolDisplay.FormatLiteral(value, quote: true);
    }

    static void EmitFormatBlock(
        StringBuilder fields,
        StringBuilder registrations,
        string rootFqn,
        List<(List<string> Path, FormatMapKind Kind)> entries)
    {
        if (entries.Count == 0)
        {
            return;
        }

        fields.Append(
            """
            static readonly global::Parchment.Generated.FormatMapEntry[] _Formats =
            {

            """);
        foreach (var (path, kind) in entries)
        {
            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", global::Parchment.Generated.FormatMapKind.");
            fields.Append(kind);
            fields.Append(", ");
            EmitGetter(fields, rootFqn, path);
            fields.AppendLine("),");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterFormat(typeof({rootFqn}), _Formats);");
    }

    static void EmitStringListBlock(
        StringBuilder fields,
        StringBuilder registrations,
        string rootFqn,
        List<List<string>> entries)
    {
        if (entries.Count == 0)
        {
            return;
        }

        fields.Append(
            """
            static readonly global::Parchment.Generated.StringListMapEntry[] _StringLists =
            {

            """);
        foreach (var path in entries)
        {
            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", ");
            EmitGetter(fields, rootFqn, path);
            fields.AppendLine("),");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterStringList(typeof({rootFqn}), _StringLists);");
    }

    static void EmitEditableBlock(
        StringBuilder fields,
        StringBuilder registrations,
        string rootFqn,
        List<(List<string> Path, MemberEntry Member)> entries)
    {
        if (entries.Count == 0)
        {
            return;
        }

        fields.Append(
            """
            static readonly global::Parchment.Generated.EditableFieldMapEntry[] _Editables =
            {

            """);
        foreach (var (path, member) in entries)
        {
            fields.Append("  ");
            EmitEditableEntry(fields, rootFqn, path, member);
            fields.AppendLine(",");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterEditable(typeof({rootFqn}), _Editables);");
    }

    // For each object between an element and one of its editable members, the type extraction can
    // create there, or null where it cannot (no setter, an abstract type, no parameterless
    // constructor). Null overall when nothing on the path can be created, which leaves the plain
    // parent-routed setter in place.
    static List<string?>? CreatableParents(string elementFqn, List<string> path, Dictionary<string, TypeEntry> typesByFqn)
    {
        var parents = new List<string?>();
        var any = false;
        var currentFqn = elementFqn;
        for (var i = 0; i < path.Count - 1; i++)
        {
            MemberEntry? hop = null;
            if (typesByFqn.TryGetValue(currentFqn, out var current))
            {
                foreach (var candidate in current.Members)
                {
                    if (candidate.Name == path[i])
                    {
                        hop = candidate;
                        break;
                    }
                }
            }

            if (hop == null)
            {
                return null;
            }

            if (hop.CanCreateValue)
            {
                parents.Add(hop.TypeFullyQualifiedName);
                any = true;
            }
            else
            {
                parents.Add(null);
            }

            currentFqn = hop.TypeFullyQualifiedName;
        }

        if (any)
        {
            return parents;
        }

        return null;
    }

    static void EmitEditableEntry(StringBuilder builder, string rootFqn, List<string> path, MemberEntry member, List<string?>? creatableParents = null)
    {
        builder.Append("new(\"");
        builder.AppendJoin('.', path);
        builder.Append("\", global::Parchment.Generated.EditableFieldKind.");
        builder.Append(member.EditableKind!.Value);
        builder.Append(", typeof(");
        builder.Append(StripNullableSuffix(member.TypeFullyQualifiedName));
        builder.Append("), ");
        builder.Append(member.EditableIsNullable ? "true" : "false");
        builder.Append(", ");
        EmitGetter(builder, rootFqn, path);
        builder.Append(", ");
        if (creatableParents == null)
        {
            EmitSetter(builder, rootFqn, path, member);
        }
        else
        {
            EmitCreatingSetter(builder, rootFqn, path, member, creatableParents);
        }

        builder.Append(", ");
        EmitCanReach(builder, rootFqn, path);
        builder.Append(", ");
        builder.Append(member.EditableMultiLine ? "true" : "false");
        builder.Append(", ");
        builder.Append(ToStringLiteral(member.EditableDateFormat));
        builder.Append(')');
    }

    static void EmitCollectionBlock(
        StringBuilder fields,
        StringBuilder registrations,
        string rootFqn,
        Dictionary<string, TypeEntry> typesByFqn,
        List<(List<string> Path, MemberEntry Member)> entries)
    {
        if (entries.Count == 0)
        {
            return;
        }

        fields.Append(
            """
            static readonly global::Parchment.Generated.CollectionFieldMapEntry[] _EditableCollections =
            {

            """);
        foreach (var (path, member) in entries)
        {
            var elementFqn = member.EditableCollectionElementFqn!;
            var isArray = member.TypeFullyQualifiedName.EndsWith("[]", StringComparison.Ordinal);

            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", typeof(");
            fields.Append(elementFqn);
            fields.Append("), ");
            EmitSetter(fields, rootFqn, path, member);
            fields.Append(", ");
            EmitCanReach(fields, rootFqn, path);
            // Activator rather than `new`: an element type with `required` members has no legal
            // object-creation expression here, and extraction (like the old runtime factory) fills
            // the members from the document right after construction.
            fields.Append(", static () => global::System.Activator.CreateInstance(typeof(");
            fields.Append(elementFqn);
            fields.Append("))!, ");
            EmitElementFields(fields, elementFqn, typesByFqn);
            fields.Append(", ");
            fields.Append(isArray ? "true" : "false");
            fields.AppendLine("),");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterEditableCollections(typeof({rootFqn}), _EditableCollections);");
    }

    // The element type's editable members, getters/setters rooted at an element instance (item-relative
    // dotted paths) — the sub-entries read/written inside each repeated section item.
    static void EmitElementFields(StringBuilder builder, string elementFqn, Dictionary<string, TypeEntry> typesByFqn)
    {
        var elementEditables = new List<(List<string> Path, MemberEntry Member)>();
        WalkForMaps(
            elementFqn,
            [],
            new(StringComparer.Ordinal)
            {
                elementFqn
            },
            typesByFqn,
            [],
            [],
            [],
            elementEditables,
            [],
            [],
            null);

        builder.Append("new global::Parchment.Generated.EditableFieldMapEntry[] { ");
        foreach (var (path, member) in elementEditables)
        {
            EmitEditableEntry(builder, elementFqn, path, member, CreatableParents(elementFqn, path, typesByFqn));
            builder.Append(", ");
        }

        builder.Append('}');
    }

    // "decimal?" → "decimal" for typeof() — the runtime entry carries the underlying type
    // (nullability travels separately as IsNullable). Char compare instead of EndsWith: the
    // char overload doesn't exist on netstandard2.0.
    static string StripNullableSuffix(string fqn)
    {
        if (fqn.Length > 0 && fqn[^1] == '?')
        {
            return fqn[..^1];
        }

        return fqn;
    }

    /// <summary>
    /// Setter shape mirrors runtime <c>EditableMap.BuildSetter</c>: walk to the PARENT object
    /// (null-conditional per hop), assign on it — never on the root. Callers validate
    /// reachability first, so a null parent is a silent no-op rather than a throw.
    /// </summary>
    static void EmitSetter(StringBuilder builder, string rootFqn, List<string> path, MemberEntry member)
    {
        var cast = SetterCast(member);

        builder.Append("(o, v) => { ");
        if (path.Count == 1)
        {
            builder.Append("((");
            builder.Append(rootFqn);
            builder.Append(")o).");
            builder.Append(path[0]);
            builder.Append(" = ");
            builder.Append(cast);
            builder.Append("; }");
            return;
        }

        builder.Append("var p = ((");
        builder.Append(rootFqn);
        builder.Append(")o)");
        AppendParentPath(builder, path);
        builder.Append("; if (p != null) { p.");
        builder.Append(path[^1]);
        builder.Append(" = ");
        builder.Append(cast);
        builder.Append("; } }");
    }

    static string SetterCast(MemberEntry member)
    {
        var castFqn = member.TypeFullyQualifiedName;
        if (member.EditableIsNullable &&
            castFqn[^1] != '?')
        {
            castFqn += "?";
        }

        if (member.EditableIsNullable)
        {
            return $"({castFqn})v";
        }

        return $"({castFqn})v!";
    }

    /// <summary>
    /// The setter for a member of an editable collection's element. Extraction builds each item
    /// from nothing, so an object on the way to the member is usually null: this walks the same
    /// parent path as <see cref="EmitSetter"/> but fills a null hop it can construct. A hop it
    /// cannot construct stays null and ends the walk, which the entry's CanReach then reports.
    /// Activator rather than `new`, for the reason the element factory uses it.
    /// </summary>
    static void EmitCreatingSetter(StringBuilder builder, string rootFqn, List<string> path, MemberEntry member, List<string?> creatableParents)
    {
        builder.Append("(o, v) => { var p0 = (");
        builder.Append(rootFqn);
        builder.Append(")o; ");
        for (var i = 0; i < creatableParents.Count; i++)
        {
            var parent = $"p{i}";
            var current = $"p{i + 1}";
            builder.Append("var ");
            builder.Append(current);
            builder.Append(" = ");
            builder.Append(parent);
            builder.Append(i == 0 ? "." : "?.");
            builder.Append(path[i]);
            builder.Append("; ");

            var typeFqn = creatableParents[i];
            if (typeFqn == null)
            {
                continue;
            }

            builder.Append("if (");
            builder.Append(parent);
            builder.Append(" != null && ");
            builder.Append(current);
            builder.Append(" == null) { ");
            builder.Append(current);
            builder.Append(" = (");
            builder.Append(typeFqn);
            builder.Append(")global::System.Activator.CreateInstance(typeof(");
            builder.Append(typeFqn);
            builder.Append("))!; ");
            builder.Append(parent);
            builder.Append('.');
            builder.Append(path[i]);
            builder.Append(" = ");
            builder.Append(current);
            builder.Append("; } ");
        }

        builder.Append("if (p");
        builder.Append(creatableParents.Count);
        builder.Append(" != null) { p");
        builder.Append(creatableParents.Count);
        builder.Append('.');
        builder.Append(path[^1]);
        builder.Append(" = ");
        builder.Append(SetterCast(member));
        builder.Append("; } }");
    }

    static void EmitCanReach(StringBuilder builder, string rootFqn, List<string> path)
    {
        if (path.Count == 1)
        {
            builder.Append("static o => true");
            return;
        }

        builder.Append("o => ((");
        builder.Append(rootFqn);
        builder.Append(")o)");
        AppendParentPath(builder, path);
        builder.Append(" != null");
    }

    static void AppendParentPath(StringBuilder builder, List<string> path)
    {
        for (var i = 0; i < path.Count - 1; i++)
        {
            builder.Append(i == 0 ? "." : "?.");
            builder.Append(path[i]);
        }
    }

    static void EmitGetter(StringBuilder builder, string rootFqn, List<string> path)
    {
        builder.Append("o => ((");
        builder.Append(rootFqn);
        builder.Append(")o)");
        for (var i = 0; i < path.Count; i++)
        {
            // First segment uses `.` (after the cast the root reference is non-null when invoked
            // through TemplateStore.Render, which never passes null). Subsequent segments use `?.`
            // so a null intermediate short-circuits — matching the runtime ChainGetter semantics.
            builder.Append(i == 0 ? "." : "?.");
            builder.Append(path[i]);
        }
    }

    /// <summary>
    /// Folds a fully qualified type name into identifier characters, with no two names folding to
    /// the same text: <c>global::Sample.A_B</c> and <c>global::Sample.A.B</c> are both legal in one
    /// model graph, and each needs a field of its own.
    /// </summary>
    /// <remarks>
    /// Letters and digits pass through, <c>.</c> becomes <c>_</c>, <c>::</c> becomes <c>__</c>, and
    /// every other character — an underscore included — becomes its code between underscores
    /// (<c>_95_</c>). Reading the result back is unambiguous because an identifier cannot start
    /// with a digit: an underscore followed by a digit always opens a code, and any other
    /// underscore is a separator. Names without an underscore read the way they always did.
    /// </remarks>
    public static string ToIdentifier(string fqn)
    {
        var builder = new StringBuilder(fqn.Length);
        for (var i = 0; i < fqn.Length; i++)
        {
            var c = fqn[i];
            if (char.IsLetterOrDigit(c))
            {
                builder.Append(c);
                continue;
            }

            if (c == '.')
            {
                builder.Append('_');
                continue;
            }

            if (c == ':' &&
                i + 1 < fqn.Length &&
                fqn[i + 1] == ':')
            {
                builder.Append("__");
                i++;
                continue;
            }

            builder.Append('_');
            builder.Append(((int) c).ToString(System.Globalization.CultureInfo.InvariantCulture));
            builder.Append('_');
        }

        return builder.ToString();
    }
}
