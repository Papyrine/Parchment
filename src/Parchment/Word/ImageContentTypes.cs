/// <summary>
/// Maps the content type handed to <see cref="IOpenXmlContext.AddImagePart"/> to the image part
/// type the Open XML SDK declares for it. A type with no part type throws rather than falling
/// back to PNG: bytes stored under the wrong declared type are a picture Word cannot show.
/// </summary>
static class ImageContentTypes
{
    // Every member of the SDK's ImagePartType, keyed by the content type the SDK declares for it
    // (ImageContentTypeTests holds this list to the SDK's), followed by the aliases in common use.
    // The SDK's second tiff member, ImagePartType.Tif, declares the unregistered "image/tif"; that
    // spelling is accepted as an alias and stored under the registered "image/tiff".
    static readonly Dictionary<string, PartTypeInfo> map = new(StringComparer.OrdinalIgnoreCase)
    {
        ["image/bmp"] = ImagePartType.Bmp,
        ["image/gif"] = ImagePartType.Gif,
        ["image/png"] = ImagePartType.Png,
        ["image/tiff"] = ImagePartType.Tiff,
        ["image/x-icon"] = ImagePartType.Icon,
        ["image/x-pcx"] = ImagePartType.Pcx,
        ["image/jpeg"] = ImagePartType.Jpeg,
        ["image/jp2"] = ImagePartType.Jp2,
        ["image/x-emf"] = ImagePartType.Emf,
        ["image/x-wmf"] = ImagePartType.Wmf,
        ["image/svg+xml"] = ImagePartType.Svg,

        ["image/x-ms-bmp"] = ImagePartType.Bmp,
        ["image/x-png"] = ImagePartType.Png,
        ["image/tif"] = ImagePartType.Tiff,
        ["image/vnd.microsoft.icon"] = ImagePartType.Icon,
        ["image/ico"] = ImagePartType.Icon,
        ["image/pcx"] = ImagePartType.Pcx,
        ["image/jpg"] = ImagePartType.Jpeg,
        ["image/pjpeg"] = ImagePartType.Jpeg,
        ["image/emf"] = ImagePartType.Emf,
        ["image/wmf"] = ImagePartType.Wmf,
        ["image/svg"] = ImagePartType.Svg
    };

    public static PartTypeInfo Resolve(string contentType)
    {
        ArgumentNullException.ThrowIfNull(contentType);

        // Only the media type decides the part type: "image/png; charset=binary" is still a png.
        var mediaType = contentType;
        var parameters = mediaType.IndexOf(';');
        if (parameters >= 0)
        {
            mediaType = mediaType[..parameters];
        }

        if (map.TryGetValue(mediaType.Trim(), out var partType))
        {
            return partType;
        }

        throw new ArgumentException(
            $"Image content type '{contentType}' is not supported. Supported content types: {string.Join(", ", map.Keys)}.",
            nameof(contentType));
    }
}
