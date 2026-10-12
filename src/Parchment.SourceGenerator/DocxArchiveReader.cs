static class DocxArchiveReader
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XName paragraphName = w + "p";
    static readonly XName textName = w + "t";
    static readonly XName fallbackName = XName.Get("Fallback", "http://schemas.openxmlformats.org/markup-compatibility/2006");

    public static DocxReadResult Read(string filePath)
    {
        var paragraphs = new List<string>();
        var bodyParagraphs = new List<string>();
        var bodyFallbacks = new List<bool>();
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

            foreach (var paragraph in doc.Descendants(paragraphName))
            {
                builder.Clear();
                AppendOwnText(paragraph, builder);

                if (builder.Length > 0)
                {
                    var text = builder.ToString();
                    paragraphs.Add(text);
                    containers.Add(ContainerId(paragraph, containerIds, ref containerCount));
                    if (isBody)
                    {
                        bodyParagraphs.Add(text);
                        bodyFallbacks.Add(InFallback(paragraph));
                    }
                }
            }
        }

        return new(paragraphs, bodyParagraphs, bodyFallbacks, containers, hasRemovePersonalInformation);
    }

    // The VML copy of a text box, which mirrors the DrawingML copy in the mc:Choice beside it.
    static bool InFallback(XElement paragraph)
    {
        foreach (var ancestor in paragraph.Ancestors())
        {
            if (ancestor.Name == fallbackName)
            {
                return true;
            }
        }

        return false;
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

    // Mirrors the runtime's ParagraphText.Texts: a text box anchored in a paragraph holds
    // paragraphs of its own, which the outer Descendants walk visits separately, so their text
    // is not also the anchoring paragraph's.
    static void AppendOwnText(XElement element, StringBuilder builder)
    {
        foreach (var child in element.Elements())
        {
            if (child.Name == textName)
            {
                builder.Append(child.Value);
                continue;
            }

            if (child.Name != paragraphName)
            {
                AppendOwnText(child, builder);
            }
        }
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