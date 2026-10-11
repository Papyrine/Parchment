static class DocxArchiveReader
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static DocxReadResult Read(string filePath)
    {
        var paragraphs = new List<string>();
        var bodyParagraphs = new List<string>();
        var containers = new List<int>();
        var containerCount = 0;
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

            // Zip entry order is not guaranteed body-first, so track the body part explicitly —
            // editable-field validation is body-scoped.
            var isBody = IsBodyPart(entry.FullName);

            using var partStream = entry.Open();
            var doc = XDocument.Load(partStream);

            // Numbered in order of first appearance and never reset between parts, so the ids are
            // stable for a given document and a header paragraph never shares one with the body.
            var containerIds = new Dictionary<XElement, int>();

            foreach (var paragraph in doc.Descendants(w + "p"))
            {
                builder.Clear();
                foreach (var t in paragraph.Descendants(w + "t"))
                {
                    builder.Append(t.Value);
                }

                if (builder.Length > 0)
                {
                    var text = builder.ToString();
                    paragraphs.Add(text);
                    containers.Add(ContainerId(paragraph, containerIds, ref containerCount));
                    if (isBody)
                    {
                        bodyParagraphs.Add(text);
                    }
                }
            }
        }

        return new(paragraphs, bodyParagraphs, containers, hasRemovePersonalInformation);
    }

    static int ContainerId(XElement paragraph, Dictionary<XElement, int> ids, ref int count)
    {
        var parent = paragraph.Parent;
        if (parent == null)
        {
            return count++;
        }

        if (ids.TryGetValue(parent, out var id))
        {
            return id;
        }

        ids[parent] = count;
        return count++;
    }

    static bool IsBodyPart(string path) =>
        path.StartsWith("word/document.xml", StringComparison.OrdinalIgnoreCase);

    static bool LooksLikeWordPart(string path) =>
        path.StartsWith("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/footnotes.xml", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("word/endnotes.xml", StringComparison.OrdinalIgnoreCase);

    static bool IsSettingsPart(string path) =>
        path.Equals("word/settings.xml", StringComparison.OrdinalIgnoreCase);
}