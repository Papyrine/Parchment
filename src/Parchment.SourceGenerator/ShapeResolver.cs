/// <summary>
/// Validation-time counterpart to <see cref="ShapeBuilder"/>. Resolves liquid member paths
/// against a pre-baked <see cref="ModelShape"/> instead of live symbols.
/// </summary>
static class ShapeResolver
{
    public static bool TryResolve(
        ModelShape shape,
        IReadOnlyList<string> segments,
        IReadOnlyDictionary<string, string> scope,
        [NotNullWhen(true)] out string? typeFqn)
    {
        typeFqn = null;
        if (segments.Count == 0)
        {
            return false;
        }

        string currentFqn;
        int start;
        if (scope.TryGetValue(segments[0], out var scoped))
        {
            currentFqn = scoped;
            start = 1;
        }
        else
        {
            currentFqn = shape.RootTypeFullyQualifiedName;
            start = 0;
        }

        for (var i = start; i < segments.Count; i++)
        {
            var entry = FindType(shape, currentFqn);
            if (entry == null)
            {
                return false;
            }

            // The root binds through the model's own accessors whatever else it implements, so
            // only what is reached past it can be a liquid string, array or dictionary.
            if (i > 0 &&
                entry.Liquid != LiquidKind.None)
            {
                if (!TryResolveLiquidMember(entry, segments[i], out var liquidFqn))
                {
                    return false;
                }

                currentFqn = liquidFqn;
                continue;
            }

            string? matched = null;
            foreach (var member in entry.Members)
            {
                if (string.Equals(member.Name, segments[i], StringComparison.OrdinalIgnoreCase))
                {
                    matched = member.TypeFullyQualifiedName;
                    break;
                }
            }

            if (matched == null)
            {
                return false;
            }

            currentFqn = matched;
        }

        typeFqn = currentFqn;
        return true;
    }

    // Lockstep with the runtime's LiquidMembers.TryResolve: Fluid answers member access on these
    // values itself, with `size` on all three and `first` / `last` on a string or an array, and
    // compares the name ordinally. None of the type's CLR members bind. An enum resolves nothing.
    static bool TryResolveLiquidMember(TypeEntry entry, string name, [NotNullWhen(true)] out string? typeFqn)
    {
        typeFqn = null;
        if (entry.Liquid == LiquidKind.Enum)
        {
            return false;
        }

        if (name == "size")
        {
            typeFqn = "int";
            return true;
        }

        if (entry.Liquid == LiquidKind.Dictionary ||
            name is not ("first" or "last"))
        {
            return false;
        }

        if (entry.Liquid == LiquidKind.String)
        {
            typeFqn = entry.TypeFullyQualifiedName;
            return true;
        }

        // A non-generic enumerable has an element, but nothing is known about it.
        typeFqn = entry.ElementTypeFullyQualifiedName ?? "object";
        return true;
    }

    /// <summary>
    /// Whether a path <see cref="TryResolve"/> turned down stopped at an enum, and which part of
    /// the path that enum is. Asked only to word the diagnostic, so it resolves again instead of
    /// widening what every successful resolve returns.
    /// </summary>
    public static bool StopsAtEnum(
        ModelShape shape,
        IReadOnlyList<string> segments,
        IReadOnlyDictionary<string, string> scope,
        [NotNullWhen(true)] out string? enumPath)
    {
        enumPath = null;

        // The longest part of the path that resolves is where the walk stopped.
        for (var length = segments.Count - 1; length > 0; length--)
        {
            var resolved = segments.Take(length).ToList();
            if (!TryResolve(shape, resolved, scope, out var typeFqn))
            {
                continue;
            }

            if (FindType(shape, typeFqn)?.Liquid != LiquidKind.Enum)
            {
                return false;
            }

            enumPath = string.Join('.', resolved);
            return true;
        }

        return false;
    }

    public static bool TryGetElementType(ModelShape shape, string typeFqn, [NotNullWhen(true)] out string? elementFqn)
    {
        elementFqn = FindType(shape, typeFqn)?.ElementTypeFullyQualifiedName;
        return elementFqn != null;
    }

    /// <summary>
    /// Returns true when the given path (walked from the root model, honoring loop scope) ends at
    /// a member marked with <c>[ExcelsiorTable]</c>. Used by the generator to gate PARCH007 /
    /// PARCH008 diagnostics without mutating the resolver's primary return signature.
    /// </summary>
    public static bool IsExcelsiorTableMember(
        ModelShape shape,
        IReadOnlyList<string> segments,
        IReadOnlyDictionary<string, string> scope)
    {
        if (segments.Count == 0)
        {
            return false;
        }

        string currentFqn;
        int start;
        if (scope.TryGetValue(segments[0], out var scoped))
        {
            currentFqn = scoped;
            start = 1;
        }
        else
        {
            currentFqn = shape.RootTypeFullyQualifiedName;
            start = 0;
        }

        for (var i = start; i < segments.Count; i++)
        {
            var entry = FindType(shape, currentFqn);
            if (entry == null)
            {
                return false;
            }

            MemberEntry? matched = null;
            foreach (var member in entry.Members)
            {
                if (string.Equals(member.Name, segments[i], StringComparison.OrdinalIgnoreCase))
                {
                    matched = member;
                    break;
                }
            }

            if (matched == null)
            {
                return false;
            }

            if (i == segments.Count - 1)
            {
                return matched.IsExcelsiorTable;
            }

            currentFqn = matched.TypeFullyQualifiedName;
        }

        return false;
    }

    public static bool TryResolveMember(
        ModelShape shape,
        IReadOnlyList<string> segments,
        IReadOnlyDictionary<string, string> scope,
        [NotNullWhen(true)] out MemberEntry? member)
    {
        member = null;
        if (segments.Count == 0)
        {
            return false;
        }

        string currentFqn;
        int start;
        if (scope.TryGetValue(segments[0], out var scoped))
        {
            currentFqn = scoped;
            start = 1;
        }
        else
        {
            currentFqn = shape.RootTypeFullyQualifiedName;
            start = 0;
        }

        for (var i = start; i < segments.Count; i++)
        {
            var entry = FindType(shape, currentFqn);
            if (entry == null)
            {
                return false;
            }

            MemberEntry? matched = null;
            foreach (var candidate in entry.Members)
            {
                if (string.Equals(candidate.Name, segments[i], StringComparison.OrdinalIgnoreCase))
                {
                    matched = candidate;
                    break;
                }
            }

            if (matched == null)
            {
                return false;
            }

            if (i == segments.Count - 1)
            {
                member = matched;
                return true;
            }

            currentFqn = matched.TypeFullyQualifiedName;
        }

        return false;
    }

    static TypeEntry? FindType(ModelShape shape, string typeFqn)
    {
        foreach (var entry in shape.Types)
        {
            if (entry.TypeFullyQualifiedName == typeFqn)
            {
                return entry;
            }
        }

        return null;
    }
}
