static class DocxArchiveReader
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XName paragraphName = w + "p";
    static readonly XName textName = w + "t";
    static readonly XName sdtName = w + "sdt";
    static readonly XName sdtPropertiesName = w + "sdtPr";
    static readonly XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    static readonly HashSet<XName> valueControlKinds =
    [
        w + "text",
        w + "date",
        w + "dropDownList",
        w + "comboBox",
        w14 + "checkbox",
        w + "picture"
    ];

    public static DocxReadResult Read(string filePath)
    {
        var paragraphs = new List<string>();
        var bodyParagraphs = new List<string>();
        var containers = new List<int>();
        var valueControlSpans = new List<int>();
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
                var spans = isBody ? valueControlSpans : null;
                AppendOwnText(paragraph, builder, spans, bodyParagraphs.Count, InsideValueControl(paragraph));

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

        return new(paragraphs, bodyParagraphs, containers, valueControlSpans, hasRemovePersonalInformation);
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
    //
    // spans collects, for the body part, the stretches of that text sitting inside a content
    // control that holds a single typed value: an [EditableField] token there would render as a
    // control nested in it, which Word refuses to open (PARCH016).
    static void AppendOwnText(XElement element, StringBuilder builder, List<int>? spans, int paragraphIndex, bool insideValueControl)
    {
        foreach (var child in element.Elements())
        {
            if (child.Name == textName)
            {
                if (insideValueControl &&
                    spans != null &&
                    child.Value.Length > 0)
                {
                    spans.Add(paragraphIndex);
                    spans.Add(builder.Length);
                    spans.Add(builder.Length + child.Value.Length);
                }

                builder.Append(child.Value);
                continue;
            }

            if (child.Name != paragraphName)
            {
                AppendOwnText(child, builder, spans, paragraphIndex, insideValueControl || IsValueControl(child));
            }
        }
    }

    static bool InsideValueControl(XElement paragraph)
    {
        foreach (var ancestor in paragraph.Ancestors(sdtName))
        {
            if (IsValueControl(ancestor))
            {
                return true;
            }
        }

        return false;
    }

    // Lockstep with the runtime's EditableTokenValidator.ValueControlKind.
    static bool IsValueControl(XElement element)
    {
        if (element.Name != sdtName)
        {
            return false;
        }

        var properties = element.Element(sdtPropertiesName);
        if (properties == null)
        {
            return false;
        }

        foreach (var child in properties.Elements())
        {
            if (valueControlKinds.Contains(child.Name))
            {
                return true;
            }
        }

        return false;
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