/// <summary>
/// Finds the additional file a <c>[ParchmentModel]</c> attribute names.
/// </summary>
/// <remarks>
/// A path is read two ways, because both are natural to write and neither is more correct. It is
/// resolved against the directory of the file the attribute sits in — so a model beside its
/// template can say <c>report.md</c>, and one in <c>Stocktakes/</c> can say
/// <c>Templates/report.md</c>, <c>./Templates/report.md</c>, or
/// <c>../Stocktakes/Templates/report.md</c>. It is also matched against the tail of every
/// additional file's path, which is what a project relies on when its templates sit in one folder
/// at the root while its models are scattered below.
/// <para>
/// Both readings usually land on the same file, and where they land on different ones the path is
/// ambiguous rather than resolvable — that is reported instead of guessed at, since silently
/// preferring one reading would bind a model to a template its author was not looking at.
/// </para>
/// </remarks>
static class TemplatePathMatcher
{
    /// <summary>
    /// Every distinct additional file the path names. Empty when nothing matches; more than one
    /// when the two readings disagree.
    /// </summary>
    public static List<T> FindAll<T>(IEnumerable<T> files, Func<T, string> pathOf, TargetInfo target)
    {
        var matches = new List<T>();
        var seen = new List<string>();
        foreach (var file in files)
        {
            var path = pathOf(file);
            if (!MatchesRelative(path, target) &&
                !MatchesSuffix(path, target))
            {
                continue;
            }

            // The same file can satisfy both readings; that is agreement, not ambiguity.
            var normalized = Normalize(path);
            if (seen.Contains(normalized, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            seen.Add(normalized);
            matches.Add(file);
        }

        return matches;
    }

    static bool MatchesRelative(string additionalFilePath, TargetInfo target)
    {
        // An in-memory compilation has no file to be relative to.
        if (target.DeclaringDirectory == null)
        {
            return false;
        }

        var declaring = target.DeclaringDirectory.Replace('\\', '/');
        return string.Equals(
            Normalize(additionalFilePath),
            Normalize($"{declaring}/{target.TemplatePath}"),
            StringComparison.OrdinalIgnoreCase);
    }

    // A path written with ./ or ../ never matches a tail, since no real path contains those
    // segments — such a path is relative by construction and needs no special case here.
    static bool MatchesSuffix(string additionalFilePath, TargetInfo target) =>
        additionalFilePath.Replace('\\', '/')
            .EndsWith(target.TemplatePath.Replace('\\', '/'), StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Collapses <c>.</c> and <c>..</c> segments.
    /// </summary>
    /// <remarks>
    /// Done by hand rather than through <c>Path.GetFullPath</c>, which resolves a relative path
    /// against the process's current directory — ambient state a generator has no business reading,
    /// and which the analyzer rules ban for that reason.
    /// </remarks>
    public static string Normalize(string path)
    {
        var segments = new List<string>();
        foreach (var segment in path.Replace('\\', '/').Split('/'))
        {
            if (segment.Length == 0 ||
                segment == ".")
            {
                continue;
            }

            if (segment == ".." &&
                segments.Count > 0 &&
                segments[^1] != "..")
            {
                segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }
}
