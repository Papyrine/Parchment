static class DocxArchiveReader
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static DocxReadResult Read(string filePath)
    {
        var paragraphs = new List<string>();
        var hasRemovePersonalInformation = false;
        var builder = new StringBuilder();
        using var archive = ZipFile.OpenRead(filePath);
        foreach (var entry in archive.Entries)
        {
            if (IsSettingsPart(entry.FullName))
            {
                using var stream = entry.Open();
                var settings = XDocument.Load(stream);
                hasRemovePersonalInformation = settings.Root?.Element(w + "removePersonalInformation") != null;
                continue;
            }

            if (!LooksLikeWordPart(entry.FullName))
            {
                continue;
            }

            using var partStream = entry.Open();
            var doc = XDocument.Load(partStream);

            foreach (var paragraph in doc.Descendants(w + "p"))
            {
                builder.Clear();
                foreach (var t in paragraph.Descendants(w + "t"))
                {
                    builder.Append(t.Value);
                }

                if (builder.Length > 0)
                {
                    paragraphs.Add(builder.ToString());
                }
            }
        }

        return new(paragraphs, hasRemovePersonalInformation);
    }

    static bool LooksLikeWordPart(string path) =>
        path.StartsWith("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/footnotes.xml", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/endnotes.xml", StringComparison.OrdinalIgnoreCase);

    static bool IsSettingsPart(string path) =>
        path.Equals("word/settings.xml", StringComparison.OrdinalIgnoreCase);
}