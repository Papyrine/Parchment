class OpenXmlContextImpl(
    MainDocumentPart mainPart,
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
        var imagePart = MainPart.AddImagePart(partType);
        using var stream = new MemoryStream(bytes);
        imagePart.FeedData(stream);
        return MainPart.GetIdOfPart(imagePart);
    }

    public bool TryGetStyle(string styleId, out StyleType styleType) =>
        styles.TryGet(styleId, out styleType);

    public int CreateBulletNumbering() =>
        numbering.CreateBulletNumbering();

    public int CreateOrderedNumbering(NumberFormatValues numberFormat) =>
        numbering.CreateOrderedNumbering(numberFormat);
}
