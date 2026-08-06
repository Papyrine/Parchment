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
