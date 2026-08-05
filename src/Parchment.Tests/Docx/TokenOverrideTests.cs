// begin-snippet: ImageTokenAliases
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
// end-snippet
// ReSharper disable PartialTypeWithSinglePart

public partial class TokenOverrideTests
{
    #region MarkdownPropertyModel

    [ParchmentBindable]
    public partial class NoteModel
    {
        public required string Title;
        public required TokenValue Body;
    }

    #endregion

    [Test]
    public async Task MarkdownProperty()
    {
        using var stream = new MemoryStream();
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: MarkdownPropertyContent
            # {{ Title }}

            {{ Body }}
            // end-snippet
            """);

        #region MarkdownPropertyRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<NoteModel>(template);
        await store.Render(
            new NoteModel
            {
                Title = "Weekly summary",
                Body = new MarkdownToken(
                    """
                    ## Highlights

                    - Shipped the **new feature**
                    - Closed _several_ bugs
                    - Ran a code review

                    > Stay the course
                    """)
            },
            stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    #region MarkdownFilterModel

    [ParchmentBindable]
    public partial class ArticleModel
    {
        public required string Heading;
        public required string Content;
    }

    #endregion

    [Test]
    public async Task MarkdownFilter()
    {
        using var stream = new MemoryStream();
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: MarkdownFilterContent
            # {{ Heading }}

            {{ Content | markdown }}
            // end-snippet
            """);

        #region MarkdownFilterRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ArticleModel>(template);
        await store.Render(
            new ArticleModel
            {
                Heading = "Release notes",
                Content =
                    """
                    ### Bug fixes

                    - Fixed crash on **empty input**
                    - Resolved _timeout_ in batch mode
                    """
            },
            stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    #region MutateModel

    [ParchmentBindable]
    public partial class StyledModel
    {
        public required string Label;
        public required TokenValue Highlight;
    }

    #endregion

    [Test]
    public async Task MutateParagraph()
    {
        using var stream = new MemoryStream();
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: MutateContent
            {{ Label }}

            {{ Highlight }}
            // end-snippet
            """);

        #region MutateRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<StyledModel>(template);
        await store.Render(
            new StyledModel
            {
                Label = "Before",
                Highlight = new MutateToken((paragraph, _) =>
                {
                    paragraph.Append(
                        new Run(
                            new RunProperties(
                                new Bold()),
                            new Text("Custom content")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }));
                })
            }, stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    #region HtmlPropertyModel

    [ParchmentBindable]
    public partial class PostModel
    {
        public required string Title;
        public required TokenValue Body;
    }

    #endregion

    [Test]
    public async Task HtmlProperty()
    {
        using var stream = new MemoryStream();
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: HtmlPropertyContent
            # {{ Title }}

            {{ Body }}
            // end-snippet
            """);

        #region HtmlPropertyRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<PostModel>(template);
        await store.Render(
            new PostModel
            {
                Title = "Welcome",
                Body = new HtmlToken(
                    """
                    <p>Welcome to the <b>weekly digest</b>.</p>
                    <ul>
                      <li>Search performance is up</li>
                      <li>Two regressions closed</li>
                    </ul>
                    """)
            },
            stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    #region OpenXmlPropertyModel

    [ParchmentBindable]
    public partial class ReportModel
    {
        public required string Title;
        public required TokenValue Callout;
    }

    #endregion

    [Test]
    public async Task OpenXmlProperty()
    {
        using var stream = new MemoryStream();
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: OpenXmlPropertyContent
            # {{ Title }}

            {{ Callout }}
            // end-snippet
            """);

        #region OpenXmlPropertyRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<ReportModel>(template);
        await store.Render(
            new ReportModel
            {
                Title = "Status",
                Callout = new OpenXmlToken(_ =>
                [
                    new Paragraph(
                        new Run(
                            new RunProperties(
                                new Color
                                {
                                    Val = "C00000"
                                },
                                new Bold()),
                            new Text("Critical: review required")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }))
                ])
            },
            stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task BulletListFilter()
    {
        #region BulletListFilterContent

        using var template = DocxTemplateBuilder.Build(
            """
            Tags:

            {{ Tags | bullet_list }}
            """);

        #endregion

        #region BulletListFilterRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);
        using var stream = new MemoryStream();
        await store.Render(SampleData.Invoice(), stream);

        #endregion

        stream.Position = 0;
        await Verify(stream, "docx");
    }

    #region ImageTokenModel

    [ParchmentBindable]
    public partial class BrandKit
    {
        public required string Title;
        public required TokenValue Logo;
    }

    #endregion

    [Test]
    public async Task OpenXmlImage()
    {
        using var stream = new MemoryStream();
        using var template = DocxTemplateBuilder.Build(
            """
            // begin-snippet: ImageTokenContent
            # {{ Title }}

            {{ Logo }}
            // end-snippet
            """);

        // 1x1 transparent PNG — stand-in for whatever bytes the model is carrying.
        var imageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII=");

        #region ImageTokenRender

        var store = new TemplateStore();
        store.RegisterDocxTemplate<BrandKit>(template);
        await store.Render(
            new BrandKit
            {
                Title = "Brand kit",
                Logo = new OpenXmlToken(context =>
                {
                    var relId = context.AddImagePart(imageBytes, "image/png");

                    // Word measures images in EMUs (English Metric Units): 914400 per inch.
                    const long widthEmu = 914400L;
                    const long heightEmu = 914400L;

                    var inline = new DW.Inline(
                        new DW.Extent
                        {
                            Cx = widthEmu,
                            Cy = heightEmu
                        },
                        new DW.DocProperties
                        {
                            Id = 1U,
                            Name = "Logo"
                        },
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties
                                        {
                                            Id = 0U,
                                            Name = "logo.png"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip
                                        {
                                            Embed = relId
                                        },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset
                                            {
                                                X = 0L,
                                                Y = 0L
                                            },
                                            new A.Extents
                                            {
                                                Cx = widthEmu,
                                                Cy = heightEmu
                                            }),
                                        new A.PresetGeometry(new A.AdjustValueList())
                                        {
                                            Preset = A.ShapeTypeValues.Rectangle
                                        })))
                            {
                                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                            }));

                    return [new Paragraph(new Run(new Drawing(inline)))];
                })
            },
            stream);

        #endregion

        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        var main = doc.MainDocumentPart!;
        await Assert.That(main.Document!.Body!.Descendants<Drawing>().Count()).IsEqualTo(1);
        await Assert.That(main.ImageParts.Any()).IsTrue();
    }
}
