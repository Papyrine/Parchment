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
/// <c>o => ((global::NS.Root)o).First?.Second?.Third</c>. Works for reference-type chains and
/// <c>Nullable&lt;T&gt;</c> intermediates; non-nullable value-type intermediates are an
/// unsupported case (rare in real binding models — value types are typically leaves).
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
        var excelsiorEntries = new List<(List<string> Path, string ElementFqn, string? HeadingParagraphStyle, string? BodyParagraphStyle)>();
        var formatEntries = new List<(List<string> Path, FormatMapKind Kind)>();
        var stringListEntries = new List<List<string>>();
        var editableEntries = new List<(List<string> Path, MemberEntry Member)>();

        // Fluid accessors: one block per non-empty type in the shape (System types & enums end
        // up with 0 members and are skipped).
        foreach (var type in shape.Types)
        {
            if (type.Members.Count == 0)
            {
                continue;
            }

            fluidBlocks.Add(("_Accessors_" + Sanitize(type.TypeFullyQualifiedName), type));
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
            editableEntries);

        if (fluidBlocks.Count == 0 &&
            excelsiorEntries.Count == 0 &&
            formatEntries.Count == 0 &&
            stringListEntries.Count == 0 &&
            editableEntries.Count == 0)
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

        fields.TrimTrailingNewlines();
        registrations.TrimTrailingNewlines();

        return new(fields.ToString(), registrations.ToString());
    }

    static void WalkForMaps(
        string currentTypeFqn,
        List<string> path,
        HashSet<string> visited,
        Dictionary<string, TypeEntry> typesByFqn,
        List<(List<string>, string, string?, string?)> excelsior,
        List<(List<string>, FormatMapKind)> formats,
        List<List<string>> stringLists,
        List<(List<string>, MemberEntry)> editables)
    {
        if (!typesByFqn.TryGetValue(currentTypeFqn, out var typeEntry))
        {
            return;
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
                if (typesByFqn.TryGetValue(member.TypeFullyQualifiedName, out var memberType) &&
                    memberType.ElementTypeFullyQualifiedName != null)
                {
                    excelsior.Add((nextPath, memberType.ElementTypeFullyQualifiedName, member.ExcelsiorHeadingParagraphStyle, member.ExcelsiorBodyParagraphStyle));
                }

                continue;
            }

            if (member.IsStringList)
            {
                stringLists.Add(nextPath);
                continue;
            }

            // A member that is also [EditableField] is claimed by the editable map (rich-content
            // control that extracts back to HTML), not the read-only format path — mirrors runtime
            // FormatMap.Build.
            if (member.IsHtml && !member.IsEditable)
            {
                formats.Add((nextPath, FormatMapKind.Html));
                continue;
            }

            if (member.IsMarkdown && !member.IsEditable)
            {
                formats.Add((nextPath, FormatMapKind.Markdown));
                continue;
            }

            if (member.IsEditable)
            {
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

            WalkForMaps(member.TypeFullyQualifiedName, nextPath, visited, typesByFqn, excelsior, formats, stringLists, editables);
            visited.Remove(member.TypeFullyQualifiedName);
        }
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
                static readonly KeyValuePair<string, global::Fluid.IMemberAccessor>[] {{fieldName}} =
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
        List<(List<string> Path, string ElementFqn, string? HeadingParagraphStyle, string? BodyParagraphStyle)> entries)
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
        foreach (var (path, elementFqn, headingParagraphStyle, bodyParagraphStyle) in entries)
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
            fields.AppendLine("),");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterExcelsiorTable(typeof({rootFqn}), _ExcelsiorTables);");
    }

    static string ToStringLiteral(string? value) =>
        value == null ? "null" : SymbolDisplay.FormatLiteral(value, quote: true);

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
            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", global::Parchment.Generated.EditableFieldKind.");
            fields.Append(member.EditableKind!.Value);
            fields.Append(", typeof(");
            fields.Append(StripNullableSuffix(member.TypeFullyQualifiedName));
            fields.Append("), ");
            fields.Append(member.EditableIsNullable ? "true" : "false");
            fields.Append(", ");
            EmitGetter(fields, rootFqn, path);
            fields.Append(", ");
            EmitSetter(fields, rootFqn, path, member);
            fields.Append(", ");
            EmitCanReach(fields, rootFqn, path);
            fields.Append(", ");
            fields.Append(member.EditableMultiLine ? "true" : "false");
            fields.Append(", ");
            fields.Append(ToStringLiteral(member.EditableDateFormat));
            fields.AppendLine("),");
        }

        fields.Append(
            """
            };

            """);

        registrations.AppendLine($"  global::Parchment.Generated.GeneratedRegistration.RegisterEditable(typeof({rootFqn}), _Editables);");
    }

    // "decimal?" → "decimal" for typeof() — the runtime entry carries the underlying type
    // (nullability travels separately as IsNullable). Char compare instead of EndsWith: the
    // char overload doesn't exist on netstandard2.0.
    static string StripNullableSuffix(string fqn) =>
        fqn.Length > 0 && fqn[^1] == '?' ? fqn[..^1] : fqn;

    /// <summary>
    /// Setter shape mirrors runtime <c>EditableMap.BuildSetter</c>: walk to the PARENT object
    /// (null-conditional per hop), assign on it — never on the root. Callers validate
    /// reachability first, so a null parent is a silent no-op rather than a throw.
    /// </summary>
    static void EmitSetter(StringBuilder builder, string rootFqn, List<string> path, MemberEntry member)
    {
        var castFqn = member.TypeFullyQualifiedName;
        if (member.EditableIsNullable &&
            castFqn[^1] != '?')
        {
            castFqn += "?";
        }

        var cast = member.EditableIsNullable ? $"({castFqn})v" : $"({castFqn})v!";

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

    static string Sanitize(string fqn)
    {
        var builder = new StringBuilder(fqn.Length);
        foreach (var c in fqn)
        {
            if (char.IsLetterOrDigit(c))
            {
                builder.Append(c);
            }
            else
            {
                builder.Append('_');
            }
        }

        return builder.ToString();
    }
}
