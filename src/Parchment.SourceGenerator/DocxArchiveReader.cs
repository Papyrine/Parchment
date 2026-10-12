static class DocxArchiveReader
{
    static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace strict = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    static readonly XName transitionalParagraph = w + "p";
    static readonly XName transitionalText = w + "t";
    static readonly XName strictParagraph = strict + "p";
    static readonly XName strictText = strict + "t";
    static readonly XName relationshipName = XName.Get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships");
    static readonly Uri packageRoot = new("http://package/");

    const string transitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    // The parts the runtime scans for tokens besides the body (DocxCloner.EnumerateParts), by the
    // last segment of their relationship type.
    static readonly string[] tokenPartTypes = ["header", "footer", "footnotes", "endnotes"];

    public static DocxReadResult Read(string filePath)
    {
        var paragraphs = new List<string>();
        var bodyParagraphs = new List<string>();
        var containers = new List<int>();
        var containerCount = 0;
        var hasRemovePersonalInformation = false;
        var builder = new StringBuilder();
        using var archive = ZipFile.OpenRead(filePath);

        // Parts are found the way the package defines them, as the runtime does: the package
        // relationships name the main part, and its own relationships name the rest. A part's
        // file name is a convention only, so word/document2.xml is as valid as word/document.xml.
        var mainPath = MainPartPath(archive);
        var mainEntry = FindEntry(archive, mainPath);
        if (mainEntry == null)
        {
            throw new InvalidDataException($"The main document part '{mainPath}' is missing from the package.");
        }

        var relationships = ReadRelationships(archive, mainPath);
        ReadPart(mainEntry, true);

        // The order of DocxCloner.EnumerateParts: body, headers, footers, footnotes, endnotes.
        var read = new List<string> {mainPath};
        foreach (var type in tokenPartTypes)
        {
            foreach (var relationship in relationships)
            {
                if (relationship.Type != type ||
                    Contains(read, relationship.Path))
                {
                    continue;
                }

                read.Add(relationship.Path);
                if (FindEntry(archive, relationship.Path) is { } entry)
                {
                    ReadPart(entry, false);
                }
            }
        }

        foreach (var relationship in relationships)
        {
            if (relationship.Type != "settings" ||
                FindEntry(archive, relationship.Path) is not { } entry)
            {
                continue;
            }

            using var stream = entry.Open();
            var root = XDocument.Load(stream).Root;
            hasRemovePersonalInformation = root?.Element(root.Name.Namespace + "removePersonalInformation") != null;
            break;
        }

        return new(paragraphs, bodyParagraphs, containers, hasRemovePersonalInformation);

        void ReadPart(ZipArchiveEntry entry, bool isBody)
        {
            using var partStream = entry.Open();
            var doc = XDocument.Load(partStream);
            var root = doc.Root!;
            var isStrict = root.Name.Namespace == strict;
            if (isBody &&
                (root.Name.LocalName != "document" ||
                 (!isStrict && root.Name.Namespace != w)))
            {
                throw new InvalidDataException($"The main part '{entry.FullName}' is not a Word document: its root element is '{root.Name.LocalName}' in namespace '{root.Name.NamespaceName}'.");
            }

            // A part is written in one namespace throughout: transitional, or Strict Open XML,
            // which the runtime's SDK reads as the same elements.
            var paragraphName = transitionalParagraph;
            var textName = transitionalText;
            if (isStrict)
            {
                paragraphName = strictParagraph;
                textName = strictText;
            }

            // Numbered in order of first appearance and never reset between parts, so the ids are
            // stable for a given document and a header paragraph never shares one with the body.
            var containerIds = new Dictionary<XElement, int>();

            foreach (var paragraph in doc.Descendants(paragraphName))
            {
                builder.Clear();
                AppendOwnText(paragraph, builder, paragraphName, textName);

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
    }

    static string MainPartPath(ZipArchive archive)
    {
        foreach (var relationship in ReadRelationships(archive, ""))
        {
            if (relationship.Type == "officeDocument")
            {
                return relationship.Path;
            }
        }

        throw new InvalidDataException("The file is not a Word package: it has no main document part.");
    }

    // The internal relationships of a part ("" for the package itself), each reduced to the last
    // segment of its type and the zip entry name its target resolves to.
    static List<(string Type, string Path)> ReadRelationships(ZipArchive archive, string sourcePath)
    {
        var relationships = new List<(string Type, string Path)>();
        var slash = sourcePath.LastIndexOf('/') + 1;
        var entry = FindEntry(archive, $"{sourcePath.Substring(0, slash)}_rels/{sourcePath.Substring(slash)}.rels");
        if (entry == null)
        {
            return relationships;
        }

        using var stream = entry.Open();
        var root = XDocument.Load(stream).Root;
        if (root == null)
        {
            return relationships;
        }

        // Targets are relative to the source part, may climb with "..", and may be rooted.
        var baseUri = new Uri(packageRoot, sourcePath);
        foreach (var element in root.Elements(relationshipName))
        {
            var target = element.Attribute("Target")?.Value;
            var type = TypeSuffix(element.Attribute("Type")?.Value);
            if (type == null ||
                target == null ||
                string.Equals(element.Attribute("TargetMode")?.Value, "External", StringComparison.OrdinalIgnoreCase) ||
                !Uri.TryCreate(baseUri, target, out var resolved))
            {
                continue;
            }

            relationships.Add((type, Uri.UnescapeDataString(resolved.AbsolutePath).TrimStart('/')));
        }

        return relationships;
    }

    // Transitional and Strict packages name the same relationship under different prefixes.
    static string? TypeSuffix(string? type)
    {
        if (type == null)
        {
            return null;
        }

        if (type.StartsWith(transitionalRelationships, StringComparison.Ordinal))
        {
            return type.Substring(transitionalRelationships.Length);
        }

        if (type.StartsWith(strictRelationships, StringComparison.Ordinal))
        {
            return type.Substring(strictRelationships.Length);
        }

        return null;
    }

    // Part names compare case-insensitively, which ZipArchive.GetEntry does not.
    static ZipArchiveEntry? FindEntry(ZipArchive archive, string path)
    {
        foreach (var entry in archive.Entries)
        {
            if (entry.FullName.Equals(path, StringComparison.OrdinalIgnoreCase))
            {
                return entry;
            }
        }

        return null;
    }

    static bool Contains(List<string> paths, string path)
    {
        foreach (var item in paths)
        {
            if (item.Equals(path, StringComparison.OrdinalIgnoreCase))
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
    static void AppendOwnText(XElement element, StringBuilder builder, XName paragraphName, XName textName)
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
                AppendOwnText(child, builder, paragraphName, textName);
            }
        }
    }
}
