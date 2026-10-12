/// <summary>
/// Finds the files a <c>[ParchmentModel]</c> type binds to by convention.
/// </summary>
/// <remarks>
/// A model named <c>Invoice</c> binds the AdditionalFile named <c>Invoice.parchment.docx</c> or
/// <c>Invoice.parchment.md</c>, wherever it sits in the project. A markdown template's style source
/// is <c>Invoice.parchment.dotx</c> when the project carries one; otherwise the nearest
/// <c>parchment.dotx</c> found walking up the directory tree from the template. There is no path to
/// write, so there is no path to get wrong — a missing or doubled file is reported against the type
/// name.
/// <para>
/// The <c>.parchment</c> marker is what lets the package discover templates on its own: a glob for
/// <c>*.parchment.md</c> cannot mistake a readme or a design note for a template, so the files need
/// no csproj entry. It is required rather than optional — an optional marker would mean two
/// conventions to know, and a file that binds under one name and not another.
/// </para>
/// </remarks>
static class TemplateConvention
{
    /// <summary>
    /// The infix that marks a file as Parchment's, and the stem of the shared style document.
    /// </summary>
    public const string Marker = "parchment";

    const string markerSuffix = $".{Marker}";

    /// <summary>A marked markdown template — <c>Invoice.parchment.md</c>.</summary>
    public static bool IsMarkdown(string path) =>
        IsMarked(path, ".md");

    /// <summary>A marked docx template — <c>Invoice.parchment.docx</c>.</summary>
    public static bool IsDocx(string path) =>
        IsMarked(path, ".docx");

    /// <summary>
    /// A style document — a marked <c>.dotx</c>, or the shared <c>parchment.dotx</c>.
    /// </summary>
    public static bool IsStyleDoc(string path) =>
        IsMarked(path, ".dotx") ||
        IsSharedStyleDoc(path);

    static bool IsMarked(string path, string extension) =>
        path.EndsWith($"{markerSuffix}{extension}", StringComparison.OrdinalIgnoreCase);

    static bool IsSharedStyleDoc(string path) =>
        string.Equals(Path.GetFileName(path), $"{Marker}.dotx", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Every file the model's type name marks as its own — <c>Invoice.parchment.md</c> for
    /// <c>Invoice</c>. An unmarked file is not a candidate, so a <c>Invoice.md</c> design note
    /// sitting in the same folder cannot bind.
    /// </summary>
    public static List<T> MatchByTypeName<T>(IEnumerable<T> files, Func<T, string> pathOf, string typeName)
    {
        var matches = new List<T>();
        foreach (var file in files)
        {
            if (string.Equals(StemOf(pathOf(file)), typeName, StringComparison.OrdinalIgnoreCase))
            {
                matches.Add(file);
            }
        }

        return matches;
    }

    /// <summary>
    /// The name a marked file binds to: <c>Invoice</c> for <c>Invoice.parchment.md</c>. Null when
    /// the file carries no marker — including the shared <c>parchment.dotx</c>, which belongs to a
    /// folder rather than to a type.
    /// </summary>
    static string? StemOf(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        if (!name.EndsWith(markerSuffix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return name[..^markerSuffix.Length];
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
            if (!IsSharedStyleDoc(dotx.Path))
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
