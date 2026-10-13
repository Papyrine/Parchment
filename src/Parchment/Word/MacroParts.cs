using System.IO.Compression;

/// <summary>
/// Takes the VBA project out of a package that is being turned into a plain document.
/// </summary>
/// <remarks>
/// Retyping a macro-enabled source (<c>.docm</c>/<c>.dotm</c>) changes the main part's content type
/// and nothing else, so on its own it leaves a document-typed package with <c>vbaProject.bin</c>
/// still inside: macros in a file the caller saves as <c>.docx</c>, which Word and security
/// tooling both treat as suspect. The output is always a plain document, so the macros go.
/// </remarks>
static class MacroParts
{
    static readonly XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    const string contentTypesEntry = "[Content_Types].xml";

    /// <summary>
    /// Deletes the VBA project and the customizations bound to it.
    /// </summary>
    /// <remarks>
    /// Deleting a part takes the parts only it refers to with it: the VBA project's supplementary
    /// data and signatures, and the toolbars attached to the key and menu customizations, which
    /// call the macros being removed.
    /// </remarks>
    /// <returns>Whether anything was deleted.</returns>
    public static bool Remove(MainDocumentPart mainPart)
    {
        var removed = false;
        if (mainPart.VbaProjectPart is { } vbaProject)
        {
            mainPart.DeletePart(vbaProject);
            removed = true;
        }

        if (mainPart.CustomizationPart is { } customization)
        {
            mainPart.DeletePart(customization);
            removed = true;
        }

        return removed;
    }

    /// <summary>
    /// Drops the extension defaults that still name a macro content type from a saved package.
    /// </summary>
    /// <remarks>
    /// A content type is declared per part or as the default for an extension, and deleting a part
    /// or retyping the package removes only the former. Word declares the VBA project as the
    /// default for <c>bin</c>, so the declaration outlives the part and the package still
    /// advertises a content type it holds nothing of. The packaging API has no way to remove a
    /// default, hence the rewrite of the one entry. A default some remaining entry still relies on
    /// is left alone: removing it would leave that entry with no content type at all.
    /// </remarks>
    public static void RemoveContentTypeDefaults(MemoryStream stream)
    {
        stream.Position = 0;
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            Rewrite(zip);
        }

        stream.Position = 0;
    }

    static void Rewrite(ZipArchive zip)
    {
        var entry = zip.GetEntry(contentTypesEntry);
        if (entry == null)
        {
            return;
        }

        XDocument document;
        using (var read = entry.Open())
        {
            document = XDocument.Load(read);
        }

        var root = document.Root!;
        var overridden = root
            .Elements(contentTypes + "Override")
            .Select(_ => (string?) _.Attribute("PartName"))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var stale = root
            .Elements(contentTypes + "Default")
            .Where(_ => IsMacroContentType((string?) _.Attribute("ContentType")) &&
                        !IsReliedOn((string?) _.Attribute("Extension"), zip, overridden))
            .ToList();
        if (stale.Count == 0)
        {
            return;
        }

        foreach (var element in stale)
        {
            element.Remove();
        }

        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false)
        };
        using var write = entry.Open();
        write.SetLength(0);
        using var writer = XmlWriter.Create(write, settings);
        document.Save(writer);
    }

    static bool IsMacroContentType(string? contentType) =>
        contentType != null &&
        (contentType.StartsWith("application/vnd.ms-office.vbaProject", StringComparison.OrdinalIgnoreCase) ||
         contentType.Equals("application/vnd.ms-word.vbaData+xml", StringComparison.OrdinalIgnoreCase) ||
         contentType.Equals("application/vnd.ms-word.document.macroEnabled.main+xml", StringComparison.OrdinalIgnoreCase) ||
         contentType.Equals("application/vnd.ms-word.template.macroEnabledTemplate.main+xml", StringComparison.OrdinalIgnoreCase) ||
         contentType.Equals("application/vnd.ms-word.keyMapCustomizations+xml", StringComparison.OrdinalIgnoreCase) ||
         contentType.Equals("application/vnd.ms-word.attachedToolbars", StringComparison.OrdinalIgnoreCase));

    static bool IsReliedOn(string? extension, ZipArchive zip, HashSet<string?> overridden)
    {
        var suffix = "." + extension;
        return zip.Entries.Any(_ => _.FullName != contentTypesEntry &&
                                    _.FullName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
                                    !overridden.Contains("/" + _.FullName));
    }
}
