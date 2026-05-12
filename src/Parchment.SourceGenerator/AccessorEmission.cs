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
        var excelsiorEntries = new List<(List<string> Path, string ElementFqn)>();
        var formatEntries = new List<(List<string> Path, FormatMapKind Kind)>();
        var stringListEntries = new List<List<string>>();

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
            stringListEntries);

        if (fluidBlocks.Count == 0 &&
            excelsiorEntries.Count == 0 &&
            formatEntries.Count == 0 &&
            stringListEntries.Count == 0)
        {
            return null;
        }

        var fields = new StringBuilder();
        var registrations = new StringBuilder();

        EmitFluidBlocks(fields, registrations, fluidBlocks);
        EmitExcelsiorBlock(fields, registrations, rootFqn, excelsiorEntries);
        EmitFormatBlock(fields, registrations, rootFqn, formatEntries);
        EmitStringListBlock(fields, registrations, rootFqn, stringListEntries);

        return new(
            fields.ToString().TrimEnd('\r', '\n'),
            registrations.ToString().TrimEnd('\r', '\n'));
    }

    static void WalkForMaps(
        string currentTypeFqn,
        List<string> path,
        HashSet<string> visited,
        Dictionary<string, TypeEntry> typesByFqn,
        List<(List<string>, string)> excelsior,
        List<(List<string>, FormatMapKind)> formats,
        List<List<string>> stringLists)
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

            var nextPath = new List<string>(path) { member.Name };

            if (member.IsExcelsiorTable)
            {
                if (typesByFqn.TryGetValue(member.TypeFullyQualifiedName, out var memberType) &&
                    memberType.ElementTypeFullyQualifiedName != null)
                {
                    excelsior.Add((nextPath, memberType.ElementTypeFullyQualifiedName));
                }

                continue;
            }

            if (member.IsStringList)
            {
                stringLists.Add(nextPath);
                continue;
            }

            if (member.IsHtml)
            {
                formats.Add((nextPath, FormatMapKind.Html));
                continue;
            }

            if (member.IsMarkdown)
            {
                formats.Add((nextPath, FormatMapKind.Markdown));
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

            WalkForMaps(member.TypeFullyQualifiedName, nextPath, visited, typesByFqn, excelsior, formats, stringLists);
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
            fields.Append("static readonly global::System.Collections.Generic.KeyValuePair<string, global::Fluid.IMemberAccessor>[] ");
            fields.Append(fieldName);
            fields.AppendLine(" =");
            fields.AppendLine("{");
            foreach (var member in type.Members)
            {
                fields.Append("  new(\"");
                fields.Append(member.Name);
                fields.Append("\", new global::Fluid.Accessors.DelegateAccessor(");
                if (member.IsStatic)
                {
                    // Static member: the instance argument is irrelevant. Access via the owner
                    // type directly so the emitted code is legal C# (CS0176 forbids instance-
                    // qualifying a static member).
                    fields.Append("(_, _) => ");
                    fields.Append(type.TypeFullyQualifiedName);
                    fields.Append('.');
                    fields.Append(member.Name);
                }
                else
                {
                    fields.Append("(o, _) => ((");
                    fields.Append(type.TypeFullyQualifiedName);
                    fields.Append(")o).");
                    fields.Append(member.Name);
                }

                fields.AppendLine(")),");
            }

            fields.AppendLine("};");
            fields.AppendLine();

            registrations.Append("  global::Parchment.Generated.GeneratedRegistration.RegisterFluidAccessors(typeof(");
            registrations.Append(type.TypeFullyQualifiedName);
            registrations.Append("), ");
            registrations.Append(fieldName);
            registrations.AppendLine(");");
        }
    }

    static void EmitExcelsiorBlock(
        StringBuilder fields,
        StringBuilder registrations,
        string rootFqn,
        List<(List<string> Path, string ElementFqn)> entries)
    {
        if (entries.Count == 0)
        {
            return;
        }

        fields.AppendLine("static readonly global::Parchment.Generated.ExcelsiorTableMapEntry[] _ExcelsiorTables =");
        fields.AppendLine("{");
        foreach (var (path, elementFqn) in entries)
        {
            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", typeof(");
            fields.Append(elementFqn);
            fields.Append("), ");
            EmitGetter(fields, rootFqn, path);
            fields.AppendLine("),");
        }

        fields.AppendLine("};");
        fields.AppendLine();

        registrations.Append("  global::Parchment.Generated.GeneratedRegistration.RegisterExcelsiorTable(typeof(");
        registrations.Append(rootFqn);
        registrations.AppendLine("), _ExcelsiorTables);");
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

        fields.AppendLine("static readonly global::Parchment.Generated.FormatMapEntry[] _Formats =");
        fields.AppendLine("{");
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

        fields.AppendLine("};");
        fields.AppendLine();

        registrations.Append("  global::Parchment.Generated.GeneratedRegistration.RegisterFormat(typeof(");
        registrations.Append(rootFqn);
        registrations.AppendLine("), _Formats);");
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

        fields.AppendLine("static readonly global::Parchment.Generated.StringListMapEntry[] _StringLists =");
        fields.AppendLine("{");
        foreach (var path in entries)
        {
            fields.Append("  new(\"");
            fields.AppendJoin('.', path);
            fields.Append("\", ");
            EmitGetter(fields, rootFqn, path);
            fields.AppendLine("),");
        }

        fields.AppendLine("};");
        fields.AppendLine();

        registrations.Append("  global::Parchment.Generated.GeneratedRegistration.RegisterStringList(typeof(");
        registrations.Append(rootFqn);
        registrations.AppendLine("), _StringLists);");
    }

    static void EmitGetter(StringBuilder sb, string rootFqn, List<string> path)
    {
        sb.Append("o => ((");
        sb.Append(rootFqn);
        sb.Append(")o)");
        for (var i = 0; i < path.Count; i++)
        {
            // First segment uses `.` (after the cast the root reference is non-null when invoked
            // through TemplateStore.Render, which never passes null). Subsequent segments use `?.`
            // so a null intermediate short-circuits — matching the runtime ChainGetter semantics.
            sb.Append(i == 0 ? "." : "?.");
            sb.Append(path[i]);
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
