public partial class ImageContentTypeTests
{
    [ParchmentBindable]
    public partial class Pictures
    {
        public required OpenXmlToken Images;
    }

    // The SDK's own list is the reference: a part type added to ImagePartType by an SDK update
    // fails here until it is mapped.
    [Test]
    public async Task EverySdkImagePartTypeIsMapped()
    {
        const System.Reflection.BindingFlags flags =
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static;
        var type = typeof(ImagePartType);
        var sdkTypes = type
            .GetProperties(flags)
            .Select(_ => _.GetValue(null))
            .Concat(type.GetFields(flags).Select(_ => _.GetValue(null)))
            .OfType<PartTypeInfo>()
            .ToList();

        await Assert.That(sdkTypes.Count).IsGreaterThanOrEqualTo(11);
        foreach (var sdkType in sdkTypes)
        {
            var resolved = ImageContentTypes.Resolve(sdkType.ContentType);

            // The SDK lists tiff twice; its unregistered "image/tif" is stored as "image/tiff".
            if (sdkType == ImagePartType.Tif)
            {
                await Assert.That(resolved).IsEqualTo(ImagePartType.Tiff);
                continue;
            }

            await Assert.That(resolved).IsEqualTo(sdkType);
        }
    }

    [Test]
    [Arguments("IMAGE/JPEG", "image/jpeg")]
    [Arguments("Image/Svg+Xml", "image/svg+xml")]
    [Arguments("image/jpeg; charset=binary", "image/jpeg")]
    [Arguments(" image/png ;q=1", "image/png")]
    [Arguments("image/jpg", "image/jpeg")]
    [Arguments("image/pjpeg", "image/jpeg")]
    [Arguments("image/x-png", "image/png")]
    [Arguments("image/svg", "image/svg+xml")]
    [Arguments("image/emf", "image/x-emf")]
    [Arguments("image/wmf", "image/x-wmf")]
    [Arguments("image/tif", "image/tiff")]
    [Arguments("image/x-ms-bmp", "image/bmp")]
    [Arguments("image/vnd.microsoft.icon", "image/x-icon")]
    public async Task CaseParametersAndAliases(string contentType, string expected) =>
        await Assert.That(ImageContentTypes.Resolve(contentType).ContentType).IsEqualTo(expected);

    // No SDK part type: these used to be stored in a part declared as png.
    [Test]
    [Arguments("image/webp")]
    [Arguments("image/avif")]
    [Arguments("application/octet-stream")]
    [Arguments("png")]
    [Arguments("")]
    public async Task UnsupportedThrows(string contentType)
    {
        ArgumentException? thrown = null;
        try
        {
            ImageContentTypes.Resolve(contentType);
        }
        catch (ArgumentException exception)
        {
            thrown = exception;
        }

        await Assert.That(thrown).IsNotNull();
        await Assert.That(thrown!.ParamName).IsEqualTo("contentType");
        await Assert.That(thrown.Message).Contains($"'{contentType}'");
        await Assert.That(thrown.Message).Contains("image/svg+xml");
    }

    [Test]
    public async Task PartsAreDeclaredAsTheirOwnType()
    {
        string[] contentTypes =
        [
            "image/png", "image/svg+xml", "image/x-emf", "image/x-wmf", "image/tiff", "image/bmp",
            "image/x-icon", "IMAGE/JPEG", "image/gif; charset=binary"
        ];
        var ids = new List<string>();

        using var template = DocxTemplateBuilder.Build("{{ Images }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Pictures>(template);
        using var stream = new MemoryStream();
        await store.Render(
            new Pictures
            {
                Images = new(context =>
                {
                    foreach (var contentType in contentTypes)
                    {
                        ids.Add(context.AddImagePart([1, 2, 3], contentType));
                    }

                    return [new Paragraph()];
                })
            },
            stream);

        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var declared = ids.Select(_ => main.GetPartById(_).ContentType).ToList();
        List<string> expected =
        [
            "image/png", "image/svg+xml", "image/x-emf", "image/x-wmf", "image/tiff", "image/bmp",
            "image/x-icon", "image/jpeg", "image/gif"
        ];
        await Assert.That(string.Join(' ', declared)).IsEqualTo(string.Join(' ', expected));
    }

    [Test]
    public async Task UnsupportedTypeFailsTheRender()
    {
        using var template = DocxTemplateBuilder.Build("{{ Images }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Pictures>(template);
        using var stream = new MemoryStream();

        Exception? thrown = null;
        try
        {
            await store.Render(
                new Pictures
                {
                    Images = new(context =>
                    {
                        context.AddImagePart([1, 2, 3], "image/webp");
                        return [new Paragraph()];
                    })
                },
                stream);
        }
        catch (Exception exception)
        {
            thrown = exception;
        }

        await Assert.That(thrown).IsNotNull();
        await Assert.That(thrown!.ToString()).Contains("Image content type 'image/webp' is not supported");
    }
}
