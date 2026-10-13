using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

// ReSharper disable PartialTypeWithSinglePart
/// <summary>
/// A loop iteration and the after half of a split paragraph are copies of template content, and a
/// copy arrives carrying the ids of what it was copied from. A drawing's <c>wp:docPr</c> id, a
/// content control's <c>w:id</c> and a tracked change's <c>w:id</c> only have to be unique, so the
/// copy is renumbered. A comment is one note with one range, so it stays on its first copy.
/// </summary>
public partial class CloneIdTests
{
    const string namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
        "xmlns:v=\"urn:schemas-microsoft-com:vml\"";

    const string tracked = "w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"";

    [AttributeUsage(AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    public class Group
    {
        public required string Name;
        public bool Shown = true;
        public List<string> Items = [];
    }

    [ParchmentBindable]
    public partial class Doc
    {
        public List<string> Items = ["a", "b", "c"];

        public List<Group> Groups =
        [
            new()
            {
                Name = "g1",
                Items = ["a", "b"]
            },
            new()
            {
                Name = "g2",
                Items = ["c", "d"]
            }
        ];

        [Html]
        public string Block = "<p>one</p><p>two</p>";

        [Html]
        public string Inline = "<b>bold</b>";
    }

    public class Row
    {
        [EditableField]
        public string? Name { get; set; }
    }

    [ParchmentBindable]
    public partial class Sheet
    {
        [EditableField]
        public required List<Row> Rows { get; set; }
    }

    [Test]
    public async Task LoopGivesEachDrawingItsOwnId()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(Drawing(1) + Run("{{ item }}")),
                P(Run("{% endfor %}"))
            ]);

        using var doc = await Render(template, new());

        // The first copy is the template's own; the later ones are numbered above it.
        await Assert.That(Join(DrawingIds(doc))).IsEqualTo("1,2,3");
        await AssertValid(doc);
    }

    [Test]
    public async Task LoopGivesEachContentControlItsOwnId()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(Control(5, Run("{{ item }}"))),
                new SdtBlock($"<w:sdt {namespaces}><w:sdtPr><w:id w:val=\"-6\"/></w:sdtPr><w:sdtContent><w:p>{Run("static")}</w:p></w:sdtContent></w:sdt>"),
                P(Run("{% endfor %}"))
            ]);

        using var doc = await Render(template, new());

        var ids = ControlIds(doc);
        await Assert.That(ids.Count).IsEqualTo(6);
        await Assert.That(ids.Distinct().Count()).IsEqualTo(6);
        await Assert.That(Join(ids.Take(2))).IsEqualTo("5,-6");
        await AssertValid(doc);
    }

    [Test]
    public async Task LoopGivesEachTrackedChangeItsOwnId()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(
                    $"<w:pPr><w:rPr><w:ins w:id=\"1\" {tracked}/></w:rPr><w:pPrChange w:id=\"2\" {tracked}><w:pPr/></w:pPrChange></w:pPr>" +
                    $"<w:ins w:id=\"3\" {tracked}>{Run("{{ item }}")}</w:ins>" +
                    $"<w:del w:id=\"4\" {tracked}><w:r><w:delText>gone</w:delText></w:r></w:del>" +
                    $"<w:r><w:rPr><w:b/><w:rPrChange w:id=\"5\" {tracked}><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r>"),
                P(Run("{% endfor %}"))
            ]);

        using var doc = await Render(template, new());

        var ids = RevisionIds(doc);
        await Assert.That(ids.Count).IsEqualTo(15);
        await Assert.That(ids.Distinct().Count()).IsEqualTo(15);
        await Assert.That(Join(ids.Take(5))).IsEqualTo("1,2,3,4,5");
        await AssertValid(doc);
    }

    [Test]
    public async Task NestedLoopsGiveEveryCopyItsOwnIds()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for group in Groups %}")),
                P(Drawing(1) + Run("{{ group.Name }}")),
                P(Run("{% for item in group.Items %}")),
                P(Drawing(2) + Control(5, Run("{{ item }}"))),
                P(Run("{% endfor %}")),
                P(Run("{% endfor %}"))
            ]);

        using var doc = await Render(template, new());

        var drawings = DrawingIds(doc);
        await Assert.That(drawings.Count).IsEqualTo(6);
        await Assert.That(drawings.Distinct().Count()).IsEqualTo(6);
        var controls = ControlIds(doc);
        await Assert.That(controls.Count).IsEqualTo(4);
        await Assert.That(controls.Distinct().Count()).IsEqualTo(4);
        await AssertValid(doc);
    }

    [Test]
    public async Task LoopInsideATableCellGivesEachCopyItsOwnIds()
    {
        var table = new Table(
            $"<w:tbl {namespaces}><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>" +
            $"<w:p>{Run("{% for item in Items %}")}</w:p><w:p>{TextBox(4)}{Run("{{ item }}")}</w:p><w:p>{Run("{% endfor %}")}</w:p>" +
            "</w:tc></w:tr></w:tbl>");
        using var template = BuildDocx([table, P(Run("after"))]);

        using var doc = await Render(template, new());

        // The drawing sits inside mc:AlternateContent, which is walked like anything else.
        await Assert.That(Join(DrawingIds(doc))).IsEqualTo("4,5,6");
        await AssertValid(doc);
    }

    [Test]
    public async Task IdsAreUniqueAcrossParts()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(Drawing(1) + Control(11, Run("{{ item }}"))),
                P(Run("{% endfor %}"))
            ],
            [
                P(Run("{% for item in Items %}")),
                P(Drawing(2) + Control(12, Run("{{ item }}"))),
                P(Run("{% endfor %}"))
            ]);

        using var doc = await Render(template, new());

        var drawings = DrawingIds(doc);
        await Assert.That(drawings.Count).IsEqualTo(6);
        await Assert.That(drawings.Distinct().Count()).IsEqualTo(6);
        var controls = ControlIds(doc);
        await Assert.That(controls.Count).IsEqualTo(6);
        await Assert.That(controls.Distinct().Count()).IsEqualTo(6);
        await AssertValid(doc);
    }

    // A header with no token is never read by a render, and what it holds still has to be stepped
    // past: the floor comes from every part at registration.
    [Test]
    public async Task NumbersAboveATokenFreeHeader()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(Drawing(1) + $"<w:ins w:id=\"1\" {tracked}>{Run("{{ item }}")}</w:ins>"),
                P(Run("{% endfor %}"))
            ],
            [P(Drawing(7) + $"<w:ins w:id=\"9\" {tracked}>{Run("static")}</w:ins>")]);

        using var doc = await Render(template, new());

        await Assert.That(Join(DrawingIds(doc))).IsEqualTo("1,8,9,7");
        await Assert.That(Join(RevisionIds(doc))).IsEqualTo("1,10,11,9");
        await AssertValid(doc);
    }

    [Test]
    public async Task SplitGivesTheAfterHalfItsOwnIds()
    {
        using var template = BuildDocx(
            [
                P(
                    $"<w:pPr><w:rPr><w:ins w:id=\"1\" {tracked}/></w:rPr></w:pPr>" +
                    $"<w:ins w:id=\"3\" {tracked}>{Run("before {{ Block }} after")}</w:ins>" +
                    Drawing(9))
            ]);

        using var doc = await Render(template, new());

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(Join(body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("before ,one,two, after");

        // Both halves carry the paragraph mark's change and the insertion the token sat inside.
        var ids = RevisionIds(doc);
        await Assert.That(ids.Count).IsEqualTo(4);
        await Assert.That(ids.Distinct().Count()).IsEqualTo(4);
        await Assert.That(Join(ids.Take(2))).IsEqualTo("1,3");

        // The drawing after the token only ever reached one half, so it is left as authored.
        await Assert.That(Join(DrawingIds(doc))).IsEqualTo("9");
        await AssertValid(doc);
    }

    [Test]
    public async Task InlineSpliceGivesTheAfterHalfItsOwnIds()
    {
        using var template = BuildDocx(
            [
                P($"<w:ins w:id=\"3\" {tracked}>{Run("before {{ Inline }} after")}</w:ins>"),
                P(Control(5, Run("before {{ Inline }} after")))
            ]);

        using var doc = await Render(template, new());

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(Join(body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("before bold after,before bold after");
        var revisions = RevisionIds(doc);
        await Assert.That(revisions.Count).IsEqualTo(2);
        await Assert.That(revisions.Distinct().Count()).IsEqualTo(2);
        await Assert.That(revisions[0]).IsEqualTo("3");
        var controls = ControlIds(doc);
        await Assert.That(controls.Count).IsEqualTo(2);
        await Assert.That(controls.Distinct().Count()).IsEqualTo(2);
        await Assert.That(controls[0]).IsEqualTo("5");
        await AssertValid(doc);
    }

    // A split inside a loop: the iteration is a copy, and its after half a copy of that.
    [Test]
    public async Task SplitInsideALoopGivesEveryHalfItsOwnIds()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P($"<w:ins w:id=\"3\" {tracked}>{Run("before {{ item | html }} after")}</w:ins>"),
                P(Run("{% endfor %}"))
            ]);
        var model = new Doc
        {
            Items = ["<p>one</p><p>two</p>", "<p>three</p><p>four</p>"]
        };

        using var doc = await Render(template, model);

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(Join(body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("before ,one,two, after,before ,three,four, after");
        var ids = RevisionIds(doc);
        await Assert.That(ids.Count).IsEqualTo(4);
        await Assert.That(ids.Distinct().Count()).IsEqualTo(4);
        await AssertValid(doc);
    }

    // An editable field numbers the controls it creates upwards from 1; a copied control counts
    // down from below zero, so neither can land on the other.
    [Test]
    public async Task EditableCollectionRowsGetTheirOwnIds()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for row in Rows %}")),
                P(Drawing(1) + Control(900, Run("label")) + Run(" {{ row.Name }}")),
                P(Run("{% endfor %}"))
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Sheet>(template);

        using var doc = await Render(
            store,
            new Sheet
            {
                Rows =
                [
                    new()
                    {
                        Name = "one"
                    },
                    new()
                    {
                        Name = "two"
                    },
                    new()
                    {
                        Name = "three"
                    }
                ]
            });

        await Assert.That(Join(DrawingIds(doc))).IsEqualTo("1,2,3");
        var controls = ControlIds(doc);
        await Assert.That(controls.Count(_ => _ == "900")).IsEqualTo(1);
        await Assert.That(controls.Distinct().Count()).IsEqualTo(controls.Count);
        await AssertValid(doc);
    }

    [Test]
    public async Task LoopKeepsACommentOnItsFirstCopy()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(Commented(Run("{{ item }}"))),
                P(Run("{% endfor %}"))
            ],
            comment: true);

        using var doc = await Render(template, new());

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(Join(body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("a,b,c");
        await AssertOneComment(body);
        var first = body.Elements<Paragraph>().First();
        await Assert.That(first.Descendants<CommentRangeStart>().Count()).IsEqualTo(1);
        await Assert.That(first.Descendants<CommentRangeEnd>().Count()).IsEqualTo(1);
        await Assert.That(first.Descendants<CommentReference>().Count()).IsEqualTo(1);

        // The run that held nothing but a dropped reference goes with it.
        await Assert.That(body.Descendants<Run>().Count()).IsEqualTo(4);
        await AssertValid(doc);
    }

    // First is the first copy to reach the output, which is not always the first iteration.
    [Test]
    public async Task CommentStaysOnTheFirstCopyThatRenders()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for group in Groups %}")),
                P(Run("{% if group.Shown %}")),
                P(Commented(Run("{{ group.Name }}"))),
                P(Run("{% endif %}")),
                P(Run("{% for item in group.Items %}")),
                P($"<w:commentRangeStart w:id=\"8\"/>{Run("{{ item }}")}<w:commentRangeEnd w:id=\"8\"/><w:r><w:commentReference w:id=\"8\"/></w:r>"),
                P(Run("{% endfor %}")),
                P(Run("{% endfor %}"))
            ],
            comment: true);
        var model = new Doc
        {
            Groups =
            [
                new()
                {
                    Name = "hidden",
                    Shown = false
                },
                new()
                {
                    Name = "shown",
                    Items = ["x", "y"]
                },
                new()
                {
                    Name = "later",
                    Items = ["z"]
                }
            ]
        };

        using var doc = await Render(template, model);

        var body = doc.MainDocumentPart!.Document!.Body!;
        await Assert.That(Join(body.Elements<Paragraph>().Select(_ => _.InnerText))).IsEqualTo("shown,x,y,later,z");
        var starts = body.Descendants<CommentRangeStart>().ToList();
        await Assert.That(Join(starts.Select(_ => $"{_.Id?.Value}:{_.Ancestors<Paragraph>().First().InnerText}"))).IsEqualTo("7:shown,8:x");
        await Assert.That(Join(body.Descendants<CommentRangeEnd>().Select(_ => _.Id?.Value ?? ""))).IsEqualTo("7,8");
        await Assert.That(Join(body.Descendants<CommentReference>().Select(_ => _.Id?.Value ?? ""))).IsEqualTo("7,8");
        await AssertValid(doc);
    }

    // The range the author drew around the paragraph still runs from its start to its end, with
    // what the token produced inside it.
    [Test]
    public async Task SplitKeepsOneCommentRangeAroundWhatWasProduced()
    {
        using var template = BuildDocx(
            [P(Commented(Run("before {{ Block }} after")))],
            comment: true);

        using var doc = await Render(template, new());

        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(Join(paragraphs.Select(_ => _.InnerText))).IsEqualTo("before ,one,two, after");
        await AssertOneComment(body);
        await Assert.That(paragraphs[0].Descendants<CommentRangeStart>().Count()).IsEqualTo(1);
        await Assert.That(paragraphs[3].Descendants<CommentRangeEnd>().Count()).IsEqualTo(1);
        await Assert.That(paragraphs[3].Descendants<CommentReference>().Count()).IsEqualTo(1);
        await AssertValid(doc);
    }

    [Test]
    public async Task InlineSpliceKeepsOneCommentRange()
    {
        using var template = BuildDocx(
            [P(Commented(Run("before {{ Inline }} after")))],
            comment: true);

        using var doc = await Render(template, new());

        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraph = body.Elements<Paragraph>().Single();
        await AssertOneComment(body);
        var order = paragraph.ChildElements.Select(
            child => child switch
            {
                CommentRangeStart => "start",
                CommentRangeEnd => "end",
                Run run when run.GetFirstChild<CommentReference>() != null => "reference",
                _ => child.InnerText
            });
        await Assert.That(Join(order)).IsEqualTo("start,before ,bold, after,end,reference");
        await AssertValid(doc);
    }

    // A range that ends before the token belongs to the before half alone.
    [Test]
    public async Task SplitLeavesACommentBeforeTheTokenInTheBeforeHalf()
    {
        using var template = BuildDocx(
            [P(Commented(Run("noted")) + Run(" then {{ Block }} after"))],
            comment: true);

        using var doc = await Render(template, new());

        var body = doc.MainDocumentPart!.Document!.Body!;
        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(Join(paragraphs.Select(_ => _.InnerText))).IsEqualTo("noted then ,one,two, after");
        await AssertOneComment(body);
        await Assert.That(paragraphs[0].Descendants<CommentRangeStart>().Count()).IsEqualTo(1);
        await Assert.That(paragraphs[0].Descendants<CommentRangeEnd>().Count()).IsEqualTo(1);
        await Assert.That(paragraphs[0].Descendants<CommentReference>().Count()).IsEqualTo(1);
        await AssertValid(doc);
    }

    // The markdown replaces the body, but a header arrives from the style source with its loops.
    [Test]
    public async Task MarkdownFlowHeaderLoopGivesEachCopyItsOwnIds()
    {
        using var styleSource = BuildDocx(
            [P(Drawing(40) + Run("replaced by the markdown"))],
            [
                P(Run("{% for item in Items %}")),
                P(Drawing(1) + Control(5, Run("{{ item }}"))),
                P(Run("{% endfor %}"))
            ]);
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>("# Title\n", styleSource);

        using var doc = await Render(store, new Doc());

        await Assert.That(Join(DrawingIds(doc))).IsEqualTo("1,2,3");
        var controls = ControlIds(doc);
        await Assert.That(controls.Count).IsEqualTo(3);
        await Assert.That(controls.Distinct().Count()).IsEqualTo(3);
        await AssertValid(doc);
    }

    [Test]
    public async Task IdsAreTheSameOnEveryRender()
    {
        using var template = BuildDocx(
            [
                P(Run("{% for item in Items %}")),
                P(Drawing(1) + Control(5, Run("{{ item }}")) + $"<w:ins w:id=\"3\" {tracked}>{Run("added")}</w:ins>"),
                P(Run("{% endfor %}"))
            ]);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);

        using var first = await Render(store, new Doc());
        using var second = await Render(store, new Doc());

        await Assert.That(Join(DrawingIds(second))).IsEqualTo(Join(DrawingIds(first)));
        await Assert.That(Join(ControlIds(second))).IsEqualTo(Join(ControlIds(first)));
        await Assert.That(Join(RevisionIds(second))).IsEqualTo(Join(RevisionIds(first)));
    }

    static async Task AssertOneComment(Body body)
    {
        await Assert.That(body.Descendants<CommentRangeStart>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<CommentRangeEnd>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<CommentReference>().Count()).IsEqualTo(1);
    }

    static async Task AssertValid(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static string Join(IEnumerable<string> values) =>
        string.Join(",", values);

    static List<string> DrawingIds(WordprocessingDocument doc) =>
        Roots(doc)
            .SelectMany(_ => _.Descendants<DW.DocProperties>())
            .Select(_ => _.Id?.InnerText ?? "")
            .ToList();

    static List<string> ControlIds(WordprocessingDocument doc) =>
        Roots(doc)
            .SelectMany(_ => _.Descendants<SdtId>())
            .Select(_ => _.Val?.InnerText ?? "")
            .ToList();

    // Every tracked change these templates hold: an inserted or deleted run, a tracked paragraph
    // mark, and a tracked change to run or paragraph properties.
    static List<string> RevisionIds(WordprocessingDocument doc) =>
        Roots(doc)
            .SelectMany(_ => _.Descendants())
            .Select(
                element => element switch
                {
                    InsertedRun change => change.Id?.Value,
                    DeletedRun change => change.Id?.Value,
                    Inserted change => change.Id?.Value,
                    RunPropertiesChange change => change.Id?.Value,
                    ParagraphPropertiesChange change => change.Id?.Value,
                    _ => null
                })
            .OfType<string>()
            .ToList();

    static IEnumerable<OpenXmlElement> Roots(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart!;
        yield return mainPart.Document!.Body!;
        foreach (var header in mainPart.HeaderParts)
        {
            yield return header.Header!;
        }
    }

    static Task<WordprocessingDocument> Render(MemoryStream template, Doc model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return Render(store, model);
    }

    static async Task<WordprocessingDocument> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    static MemoryStream BuildDocx(OpenXmlElement[] body, OpenXmlElement[]? header = null, bool comment = false)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new(new Body());
            var documentBody = mainPart.Document.Body!;
            documentBody.Append(body);

            var section = new SectionProperties();
            if (header != null)
            {
                var headerPart = mainPart.AddNewPart<HeaderPart>("rIdHeader");
                headerPart.Header = new(header);
                section.Append(
                    new HeaderReference
                    {
                        Type = HeaderFooterValues.Default,
                        Id = "rIdHeader"
                    });
            }

            if (comment)
            {
                var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>("rIdComments");
                commentsPart.Comments = new(
                    $"<w:comments {namespaces}>" +
                    $"<w:comment w:id=\"7\" w:author=\"a\"><w:p>{Run("a note")}</w:p></w:comment>" +
                    $"<w:comment w:id=\"8\" w:author=\"a\"><w:p>{Run("another note")}</w:p></w:comment>" +
                    "</w:comments>");
            }

            section.Append(
                new PageSize
                {
                    Width = 6500,
                    Height = 8000
                });
            documentBody.Append(section);
        }

        stream.Position = 0;
        return stream;
    }

    static Paragraph P(string inner) =>
        new($"<w:p {namespaces}>{inner}</w:p>");

    static string Run(string text) =>
        $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    static string Drawing(uint id) =>
        $"<w:r><w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"100\"/><wp:docPr id=\"{id}\" name=\"pic\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"/></a:graphic></wp:inline></w:drawing></w:r>";

    // A text box as Word writes one: the drawing in the choice, vml in the fallback.
    static string TextBox(uint id) =>
        "<w:r><mc:AlternateContent><mc:Choice Requires=\"wps\">" +
        $"<w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"100\"/><wp:docPr id=\"{id}\" name=\"box\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent><w:p>{Run("boxed")}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>" +
        $"</mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p>{Run("boxed")}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>";

    static string Control(int id, string content) =>
        $"<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>";

    // Comment 7 anchored around the content, as Word writes it: the range, then the reference.
    static string Commented(string content) =>
        $"<w:commentRangeStart w:id=\"7\"/>{content}<w:commentRangeEnd w:id=\"7\"/><w:r><w:commentReference w:id=\"7\"/></w:r>";
}
