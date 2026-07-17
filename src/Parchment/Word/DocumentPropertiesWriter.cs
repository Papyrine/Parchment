using DocumentFormat.OpenXml.VariantTypes;
using CustomProps = DocumentFormat.OpenXml.CustomProperties;
using ExtendedProps = DocumentFormat.OpenXml.ExtendedProperties;

/// <summary>
/// Applies a <see cref="DocumentProperties"/> to a rendered document: core properties to
/// <c>docProps/core.xml</c>, company/manager to the extended part (<c>docProps/app.xml</c>), and
/// user-defined entries to the custom part (<c>docProps/custom.xml</c>).
/// </summary>
/// <remarks>
/// Every part is <b>merged</b> rather than rewritten. A template normally arrives carrying
/// properties of its own, and replacing a part wholesale drops them silently — the trap both
/// hand-rolled stampers this replaces had to learn. Only the values the caller set are touched.
///
/// The namespace aliasing is isolated here because CustomProperties, ExtendedProperties and
/// Parchment all expose a type named <c>Properties</c>. Mirrors Excelsior's writer of the same name,
/// which builds from scratch and so can replace rather than merge.
/// </remarks>
static class DocumentPropertiesWriter
{
    // The fixed FMTID required on every custom document property by the OOXML spec.
    const string customFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    static readonly XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    static readonly XNamespace dc = "http://purl.org/dc/elements/1.1/";

    public static void Apply(WordprocessingDocument document, DocumentProperties properties)
    {
        ApplyCore(document, properties);
        ApplyExtended(document, properties);
        ApplyCustom(document, properties);
    }

    // Written as an explicit CoreFilePropertiesPart rather than through
    // OpenXmlPackage.PackageProperties: the latter is backed by the package's intrinsic
    // core-property store, which is not cloned with the rest of the parts.
    static void ApplyCore(WordprocessingDocument document, DocumentProperties properties)
    {
        if (properties is
            {
                Title: null,
                Author: null,
                Subject: null,
                Keywords: null,
                Comments: null,
                Category: null,
                Status: null,
                LastModifiedBy: null
            })
        {
            return;
        }

        var part = document.CoreFilePropertiesPart;
        XElement root;
        if (part == null)
        {
            part = document.AddCoreFilePropertiesPart();
            root = new(
                cp + "coreProperties",
                new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName));
        }
        else
        {
            root = ReadRoot(part) ??
                   new XElement(
                       cp + "coreProperties",
                       new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName),
                       new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName));
        }

        Set(root, dc + "title", properties.Title);
        Set(root, dc + "subject", properties.Subject);
        Set(root, dc + "creator", properties.Author);
        Set(root, cp + "keywords", properties.Keywords);
        Set(root, dc + "description", properties.Comments);
        Set(root, cp + "lastModifiedBy", properties.LastModifiedBy);
        Set(root, cp + "category", properties.Category);
        Set(root, cp + "contentStatus", properties.Status);

        using var stream = part.GetStream(FileMode.Create);
        new XDocument(root).Save(stream);
    }

    static XElement? ReadRoot(CoreFilePropertiesPart part)
    {
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream).Root;
        }
        catch (XmlException)
        {
            // An unreadable core part is replaced rather than failing the render — the properties
            // the caller asked for still land.
            return null;
        }
    }

    // Replaces the element's text when set, adds it when absent, and leaves it alone when the
    // caller did not ask for it.
    static void Set(XElement root, XName name, string? value)
    {
        if (value == null)
        {
            return;
        }

        var existing = root.Element(name);
        if (existing == null)
        {
            root.Add(new XElement(name, value));
            return;
        }

        existing.Value = value;
    }

    static void ApplyExtended(WordprocessingDocument document, DocumentProperties properties)
    {
        if (properties.Company == null &&
            properties.Manager == null)
        {
            return;
        }

        var part = document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart();
        var extended = part.Properties ??= new ExtendedProps.Properties();

        if (properties.Company != null)
        {
            extended.Company = new(properties.Company);
        }

        if (properties.Manager != null)
        {
            extended.Manager = new(properties.Manager);
        }
    }

    static void ApplyCustom(WordprocessingDocument document, DocumentProperties properties)
    {
        if (properties.Custom.Count == 0 &&
            properties.RemoveCustom.Count == 0)
        {
            return;
        }

        var part = document.CustomFilePropertiesPart ?? document.AddCustomFilePropertiesPart();
        var custom = part.Properties ??= new CustomProps.Properties();

        foreach (var (name, value) in properties.Custom)
        {
            var property = Find(custom, name);
            if (property == null)
            {
                property = new()
                {
                    FormatId = customFormatId,
                    Name = name
                };
                custom.AppendChild(property);
            }
            else
            {
                property.RemoveAllChildren();
            }

            property.AppendChild(ToVariant(name, value));
        }

        foreach (var name in properties.RemoveCustom)
        {
            Find(custom, name)?.Remove();
        }

        // PropertyId is 1-based and 1 is reserved, so user-defined properties start at 2. They must
        // be contiguous, so renumber after adding or removing any.
        var id = 2;
        foreach (var property in custom.Elements<CustomProps.CustomDocumentProperty>())
        {
            property.PropertyId = id;
            id++;
        }

        if (!custom.HasChildren)
        {
            document.DeletePart(part);
        }
    }

    static CustomProps.CustomDocumentProperty? Find(CustomProps.Properties custom, string name) =>
        custom.Elements<CustomProps.CustomDocumentProperty>()
            .FirstOrDefault(_ => string.Equals(_.Name?.Value, name, StringComparison.OrdinalIgnoreCase));

    static OpenXmlElement ToVariant(string name, object? value) =>
        value switch
        {
            null => new VTLPWSTR(string.Empty),
            string s => new VTLPWSTR(s),
            bool b => new VTBool(b ? "true" : "false"),
            int i => new VTInt32(i.ToString(CultureInfo.InvariantCulture)),
            long l => new VTInt64(l.ToString(CultureInfo.InvariantCulture)),
            short s => new VTInt32(((int) s).ToString(CultureInfo.InvariantCulture)),
            byte b => new VTInt32(((int) b).ToString(CultureInfo.InvariantCulture)),
            double d => new VTDouble(d.ToString(CultureInfo.InvariantCulture)),
            float f => new VTDouble(((double) f).ToString(CultureInfo.InvariantCulture)),
            decimal m => new VTDouble(m.ToString(CultureInfo.InvariantCulture)),
            DateTime time => new VTFileTime(time.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)),
            DateOnly date => new VTFileTime($"{date:yyyy-MM-dd}T00:00:00Z"),
            // Written as text rather than the clsid variant: Word's Advanced Properties dialog only
            // surfaces Text/Date/Number/Yes-No, so a string shows as an editable property whereas
            // clsid would not appear at all.
            Guid guid => new VTLPWSTR(guid.ToString()),
            // Deliberately strict rather than falling back to value.ToString(): coercing an
            // unsupported type would write something like "System.Int32[]" into the property and
            // hide the caller's mistake.
            _ => throw new ArgumentException(
                $"Custom document property '{name}' has unsupported value type '{value.GetType()}'. Supported types are string, bool, integral and floating-point numbers, DateTime, DateOnly and Guid. Convert the value to one of these before adding it.")
        };
}
