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
    static readonly XName runName = w + "r";
    static readonly XName fieldCharName = w + "fldChar";
    static readonly XName fieldCharTypeName = w + "fldCharType";
    static readonly XName formFieldDataName = w + "ffData";
    static readonly XName textInputName = w + "textInput";
    static readonly XName formFieldNameName = w + "name";
    static readonly XName valName = w + "val";
    static readonly XName deletedRunName = w + "del";
    static readonly XName moveFromRunName = w + "moveFrom";
    static readonly XName fallbackName = XName.Get("Fallback", "http://schemas.openxmlformats.org/markup-compatibility/2006");
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

    // Takes the file's bytes rather than its path: the caller embeds those same bytes, so what is
    // validated here is what ships, and the file is opened once, by the one place that knows how
    // to share it with a word processor that has it open.
    public static DocxReadResult Read(byte[] content)
    {
        var paragraphs = new List<string>();
        var bodyParagraphs = new List<string>();
        var bodyFallbacks = new List<bool>();
        var containers = new List<int>();
        var valueControlSpans = new List<int>();
        var containerCount = 0;
        var hasRemovePersonalInformation = false;
        var builder = new StringBuilder();
        using var archive = new ZipArchive(new MemoryStream(content, writable: false), ZipArchiveMode.Read);

        // Parts are found the way the package defines them, as the runtime does: the package
        // relationships name the main part, and its own relationships name the rest. A part's
        // file name is a convention only, so word/document2.xml is as valid as word/document.xml.
        var mainPath = MainPartPath(archive);
        if (mainPath == null)
        {
            throw new InvalidDataException("The file is not a Word package: it has no main document part.");
        }

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

        return new(paragraphs, bodyParagraphs, bodyFallbacks, containers, valueControlSpans, hasRemovePersonalInformation, true);

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
            // which the runtime's SDK reads as the same elements. The form field, content control
            // and text box rules below look for the transitional names only.
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

            // The runs of each FORMTEXT field the runtime rewrites into a token, keyed to the token
            // for the run that opens the field and to null for the rest. Most templates have none.
            Dictionary<XElement, string?>? fieldRuns = null;
            if (doc.Descendants(formFieldDataName).Any())
            {
                fieldRuns = [];
            }

            foreach (var paragraph in doc.Descendants(paragraphName))
            {
                if (fieldRuns != null)
                {
                    // The runtime removes a converted field's runs, and a text box anchored in one
                    // goes with it.
                    if (paragraph.Ancestors(runName).Any(fieldRuns.ContainsKey))
                    {
                        continue;
                    }

                    FindFormFields(paragraph, fieldRuns);
                }

                builder.Clear();
                var spans = isBody ? valueControlSpans : null;
                AppendOwnText(paragraph, builder, paragraphName, textName, fieldRuns, spans, bodyParagraphs.Count, InsideValueControl(paragraph));

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
    }

    /// <summary>
    /// Whether the package names a main document part that is there. A file that does not is not
    /// a Word document or template at all, which a caller may want to say in its own words.
    /// </summary>
    public static bool HasMainPart(byte[] content)
    {
        using var archive = new ZipArchive(new MemoryStream(content, writable: false), ZipArchiveMode.Read);
        return MainPartPath(archive) is { } mainPath &&
               FindEntry(archive, mainPath) != null;
    }

    static string? MainPartPath(ZipArchive archive)
    {
        foreach (var relationship in ReadRelationships(archive, ""))
        {
            if (relationship.Type == "officeDocument")
            {
                return relationship.Path;
            }
        }

        return null;
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
    //
    // spans collects, for the body part, the stretches of that text sitting inside a content
    // control that holds a single typed value: an [EditableField] token there would render as a
    // control nested in it, which Word refuses to open (PARCH016).
    static void AppendOwnText(XElement element, StringBuilder builder, XName paragraphName, XName textName, Dictionary<XElement, string?>? fieldRuns, List<int>? spans, int paragraphIndex, bool insideValueControl)
    {
        foreach (var child in element.Elements())
        {
            if (fieldRuns != null &&
                fieldRuns.TryGetValue(child, out var token))
            {
                builder.Append(token);
                continue;
            }

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
                AppendOwnText(child, builder, paragraphName, textName, fieldRuns, spans, paragraphIndex, insideValueControl || IsValueControl(child));
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

    // Mirrors the runtime's FormFields.Convert, which rewrites a legacy FORMTEXT field into a
    // {{ Name }} token run before anything scans the paragraph: the token stands where the run
    // opening the field was and every run of the field, placeholder text included, is gone. The
    // generator cannot rewrite the document, so it records the runs and AppendOwnText reads them
    // as the runtime would after the rewrite. Any change to which fields convert has to be made in
    // both.
    static void FindFormFields(XElement paragraph, Dictionary<XElement, string?> fieldRuns)
    {
        var runs = new List<XElement>();
        CollectRuns(paragraph, runs);
        for (var index = 0; index < runs.Count; index++)
        {
            var name = FieldName(runs[index]);
            if (name == null)
            {
                continue;
            }

            var end = IndexOfEnd(runs, index);
            if (end == -1)
            {
                continue;
            }

            fieldRuns[runs[index]] = $"{{{{ {name} }}}}";
            for (var scan = index + 1; scan <= end; scan++)
            {
                fieldRuns[runs[scan]] = null;
            }

            index = end;
        }
    }

    // The paragraph's own runs in document order, through hyperlinks, tracked changes, content
    // controls and the like. A run is not descended into.
    static void CollectRuns(XElement parent, List<XElement> runs)
    {
        foreach (var child in parent.Elements())
        {
            if (child.Name == runName)
            {
                runs.Add(child);
                continue;
            }

            CollectRuns(child, runs);
        }
    }

    static string? FieldName(XElement run)
    {
        // A field in a tracked deletion is content the author removed, not a placeholder.
        var parent = run.Parent?.Name;
        if (parent == deletedRunName ||
            parent == moveFromRunName)
        {
            return null;
        }

        var fieldChar = run.Element(fieldCharName);
        if (fieldChar == null ||
            FieldCharType(fieldChar) != "begin")
        {
            return null;
        }

        var data = fieldChar.Element(formFieldDataName);
        // Checkbox and dropdown fields hold values Word manages, not text to substitute.
        if (data?.Element(textInputName) == null)
        {
            return null;
        }

        var name = data.Element(formFieldNameName)?.Attribute(valName)?.Value;
        if (name == null ||
            !IsTokenName(name))
        {
            return null;
        }

        return name;
    }

    static string? FieldCharType(XElement fieldChar) =>
        fieldChar.Attribute(fieldCharTypeName)?.Value;

    static bool IsTokenName(string name)
    {
        if (name.Length == 0 ||
            !char.IsLetter(name[0]) && name[0] != '_')
        {
            return false;
        }

        foreach (var character in name)
        {
            if (!char.IsLetterOrDigit(character) &&
                character != '_' &&
                character != '.')
            {
                return false;
            }
        }

        return true;
    }

    // Fields nest, so the closing end is the one that balances the opening begin, and a field the
    // paragraph does not close is not a field to convert.
    static int IndexOfEnd(List<XElement> runs, int start)
    {
        var depth = 0;
        for (var index = start; index < runs.Count; index++)
        {
            foreach (var fieldChar in runs[index].Elements(fieldCharName))
            {
                var type = FieldCharType(fieldChar);
                if (type == "begin")
                {
                    depth++;
                    continue;
                }

                if (type != "end")
                {
                    continue;
                }

                depth--;
                if (depth == 0)
                {
                    return index;
                }
            }
        }

        return -1;
    }
}
