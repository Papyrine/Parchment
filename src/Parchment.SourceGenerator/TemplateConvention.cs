/// <summary>
/// Finds the files a <c>[ParchmentModel]</c> type binds to by convention.
/// </summary>
/// <remarks>
/// A model named <c>Invoice</c> binds the AdditionalFile named <c>Invoice.docx</c> or
/// <c>Invoice.md</c>, wherever it sits in the project. A markdown template's style source is
/// <c>Invoice.dotx</c> when the project carries one; otherwise the nearest <c>parchment.dotx</c>
/// found walking up the directory tree from the template. There is no path to write, so there is
/// no path to get wrong — a missing or doubled file is reported against the type name.
/// </remarks>
static class TemplateConvention
{
    /// <summary>
    /// Every file whose name (without extension) is the model's type name.
    /// </summary>
    public static List<T> MatchByTypeName<T>(IEnumerable<T> files, Func<T, string> pathOf, string typeName)
    {
        var matches = new List<T>();
        foreach (var file in files)
        {
            var fileName = Path.GetFileNameWithoutExtension(pathOf(file));
            if (string.Equals(fileName, typeName, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(file);
            }
        }

        return matches;
    }

    /// <summary>
    /// The <c>parchment.dotx</c> in the nearest directory at or above the template's — the shared
    /// style document a folder of templates inherits.
    /// </summary>
    public static DotxData? FindSharedStyleDoc(IEnumerable<DotxData> dotxes, string templatePath)
    {
        var templateDirectory = Normalize(GetDirectory(templatePath));
        DotxData? nearest = null;
        var nearestLength = -1;
        foreach (var dotx in dotxes)
        {
            var fileName = Path.GetFileNameWithoutExtension(dotx.Path);
            if (!string.Equals(fileName, "parchment", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var directory = Normalize(GetDirectory(dotx.Path));
            var isAncestor = string.Equals(directory, templateDirectory, StringComparison.OrdinalIgnoreCase) ||
                             templateDirectory.StartsWith($"{directory}/", StringComparison.OrdinalIgnoreCase);
            if (!isAncestor)
            {
                continue;
            }

            // The longest ancestor directory is the closest one to the template.
            if (directory.Length > nearestLength)
            {
                nearest = dotx;
                nearestLength = directory.Length;
            }
        }

        return nearest;
    }

    static string GetDirectory(string path) =>
        Path.GetDirectoryName(path) ?? "";

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
