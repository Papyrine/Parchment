/// <summary>
/// Per-render map from an id — a markdown <c>{#id}</c>, the target of a <c>#id</c> link — to the
/// name of the bookmark that stands for it.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="MarkdownBookmark.Sanitize"/> folds an id into a legal bookmark name, and folding is
/// not injective: <c>a-b</c> and <c>a_b</c> both give <c>a_b</c>, and two ids that differ only past
/// the 40th character truncate to the same name. Word keeps one bookmark per name, and compares
/// names without regard to case, so every link to the other id would land on the wrong heading.
/// The name is therefore decided here once, and the bookmark and everything that refers to it —
/// a link's <c>w:anchor</c>, a <c>PAGEREF</c> — ask for it by id.
/// </para>
/// <para>
/// An id takes its folded name when that is free, and otherwise the name with a numeric suffix,
/// shortened to stay inside the limit. Names are handed out in the order ids are first mentioned,
/// whether by a link or by the block that carries the id, because a link can come before its
/// target and has to be given the name the target will get. That order follows the template, so
/// the same template and model give the same names.
/// </para>
/// </remarks>
sealed class BookmarkNames
{
    Dictionary<string, string>? names;
    HashSet<string>? taken;
    HashSet<string>? defined;

    /// <summary>
    /// The bookmark name that <paramref name="id"/> refers to, whether or not anything carries it.
    /// </summary>
    public string For(string id)
    {
        names ??= new(StringComparer.Ordinal);
        if (names.TryGetValue(id, out var name))
        {
            return name;
        }

        name = Free(MarkdownBookmark.Sanitize(id));
        names[id] = name;
        return name;
    }

    /// <summary>
    /// The name for the bookmark of the block carrying <paramref name="id"/>, or null when an
    /// earlier block already carries that id: as in html, the first is the one a link reaches.
    /// </summary>
    public string? Define(string id)
    {
        defined ??= new(StringComparer.Ordinal);
        if (!defined.Add(id))
        {
            return null;
        }

        return For(id);
    }

    /// <summary>
    /// Claim a generated name, such as a table of contents entry's. False when it is in use.
    /// </summary>
    public bool Reserve(string name)
    {
        taken ??= new(StringComparer.OrdinalIgnoreCase);
        return taken.Add(name);
    }

    string Free(string name)
    {
        if (Reserve(name))
        {
            return name;
        }

        for (var index = 2;; index++)
        {
            var suffix = "_" + index.ToString(CultureInfo.InvariantCulture);
            var length = Math.Min(name.Length, MarkdownBookmark.MaxLength - suffix.Length);
            var candidate = name[..length] + suffix;
            if (Reserve(candidate))
            {
                return candidate;
            }
        }
    }
}
