/// <summary>
/// Reads the bookmark name from a <c>{#id}</c> generic attribute.
/// </summary>
/// <remarks>
/// A Word bookmark name is far more restricted than an html id: it must start with a letter or
/// underscore, may then carry only letters, digits and underscores, and is capped at 40 characters.
/// Rather than reject an id that html would accept, every other character is folded to an underscore
/// and the result truncated, so markdown written for both html and Word keeps working — the same
/// leniency <see cref="MarkdownStyle"/> applies to the rest of the generic-attribute syntax.
/// </remarks>
static class MarkdownBookmark
{
    // Word's documented ceiling for a bookmark name.
    const int MaxLength = 40;

    public static string? Resolve(IMarkdownObject node)
    {
        var id = node.TryGetAttributes()?.Id;
        if (string.IsNullOrEmpty(id))
        {
            return null;
        }

        return Sanitize(id);
    }

    public static string Sanitize(string id)
    {
        var builder = new StringBuilder(Math.Min(id.Length + 1, MaxLength));
        foreach (var character in id)
        {
            if (builder.Length == MaxLength)
            {
                break;
            }

            if (char.IsLetter(character) ||
                character == '_' ||
                (builder.Length > 0 && char.IsDigit(character)))
            {
                builder.Append(character);
                continue;
            }

            // A leading digit is legal in an html id but not in a bookmark name, so it is kept and
            // pushed behind the underscore that the name has to start with anyway.
            if (builder.Length == 0)
            {
                builder.Append('_');
                if (char.IsDigit(character))
                {
                    builder.Append(character);
                }

                continue;
            }

            builder.Append('_');
        }

        return builder.Length == 0 ? "_" : builder.ToString();
    }
}
