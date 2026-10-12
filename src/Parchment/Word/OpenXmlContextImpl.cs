class OpenXmlContextImpl(
    MainDocumentPart mainPart,
    OpenXmlPart part,
    WordNumberingState numbering,
    StyleSet styles,
    Paragraph? hostParagraph = null) :
    IOpenXmlContext
{
    public MainDocumentPart MainPart { get; } = mainPart;
    public int CurrentHeadingLevel { get; set; }
    public Paragraph? HostParagraph { get; } = hostParagraph;

    public string AddImagePart(byte[] bytes, string contentType)
    {
        var partType = contentType switch
        {
            "image/png" => ImagePartType.Png,
            "image/jpeg" or "image/jpg" => ImagePartType.Jpeg,
            "image/gif" => ImagePartType.Gif,
            "image/bmp" => ImagePartType.Bmp,
            "image/tiff" => ImagePartType.Tiff,
            _ => ImagePartType.Png
        };

        // To the part the token sits in, not always the main one: the id handed back is only good
        // in the part that ends up holding the drawing.
        var id = RelationshipIds.Next(part, "rIdImage");
        var imagePart = part switch
        {
            MainDocumentPart main => main.AddImagePart(partType, id),
            HeaderPart header => header.AddImagePart(partType, id),
            FooterPart footer => footer.AddImagePart(partType, id),
            FootnotesPart footnotes => footnotes.AddImagePart(partType, id),
            EndnotesPart endnotes => endnotes.AddImagePart(partType, id),
            _ => throw new InvalidOperationException($"A token cannot add an image to a {part.GetType().Name}.")
        };
        using var stream = new MemoryStream(bytes);
        imagePart.FeedData(stream);
        return id;
    }

    public bool TryGetStyle(string styleId, out StyleType styleType) =>
        styles.TryGet(styleId, out styleType);

    public int CreateBulletNumbering() =>
        numbering.CreateBulletNumbering();

    public int CreateOrderedNumbering(NumberFormatValues numberFormat) =>
        numbering.CreateOrderedNumbering(numberFormat);
}
