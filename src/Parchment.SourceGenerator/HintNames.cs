/// <summary>
/// Hint names for the generated files. A hint name is flattened from the namespace, the enclosing
/// types and the type name, which is readable but not injective: <c>A_B.C</c> and <c>A.B_C</c>
/// flatten the same, and Roslyn compares hint names without regard to case (they become file
/// names), so <c>Letter</c> and <c>letter</c> collide as well. A duplicate throws out of
/// <c>AddSource</c> and takes the whole generator's output with it, so the targets that share a
/// name are told apart here.
/// </summary>
static class HintNames
{
    public static string Flatten(TargetInfo target)
    {
        var builder = new StringBuilder();
        if (target.DeclaringNamespace != null)
        {
            builder.Append(target.DeclaringNamespace).Append('_');
        }

        foreach (var enclosing in target.EnclosingTypes)
        {
            builder.Append(enclosing.Name).Append('_');
        }

        builder.Append(target.DeclaringName);

        // The namespace is the escaped display string, and `@` is not a legal hint character.
        return builder
            .Replace("@", "")
            .Replace('.', '_')
            .ToString();
    }

    /// <summary>
    /// The targets that cannot use their flattened name, with the name each gets instead. Empty
    /// for almost every compilation, which is what lets a per-target stage combine with it and
    /// stay cached. The result depends on the set of targets only, never on their order: within a
    /// group of colliding targets the one whose fully qualified name sorts first keeps the plain
    /// name and the rest are numbered, skipping any name some other target already owns.
    /// </summary>
    public static EquatableArray<HintOverride> Disambiguate(IEnumerable<TargetInfo> targets)
    {
        var groups = new SortedDictionary<string, List<(string Model, string Flattened)>>(StringComparer.OrdinalIgnoreCase);
        var collides = false;
        foreach (var target in targets)
        {
            var flattened = Flatten(target);
            if (groups.TryGetValue(flattened, out var group))
            {
                collides = true;
            }
            else
            {
                groups[flattened] = group = [];
            }

            group.Add((target.ModelFullyQualifiedName, flattened));
        }

        if (!collides)
        {
            return EquatableArray<HintOverride>.Empty;
        }

        var taken = new HashSet<string>(groups.Keys, StringComparer.OrdinalIgnoreCase);
        var overrides = ImmutableArray.CreateBuilder<HintOverride>();
        foreach (var group in groups.Values)
        {
            group.Sort(static (left, right) => string.CompareOrdinal(left.Model, right.Model));
            var number = 1;
            for (var i = 1; i < group.Count; i++)
            {
                string candidate;
                do
                {
                    number++;
                    candidate = $"{group[i].Flattened}_{number}";
                } while (!taken.Add(candidate));

                overrides.Add(new(group[i].Model, candidate));
            }
        }

        return new(overrides.ToImmutable());
    }

    public static string Resolve(TargetInfo target, EquatableArray<HintOverride> overrides)
    {
        foreach (var item in overrides)
        {
            if (item.ModelFullyQualifiedName == target.ModelFullyQualifiedName)
            {
                return item.HintName;
            }
        }

        return Flatten(target);
    }
}
