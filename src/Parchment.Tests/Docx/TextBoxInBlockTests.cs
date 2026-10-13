// A loop iteration and a chosen branch each index the anchors of their own elements. A text box
// is anchored in a run of one paragraph and holds paragraphs of its own, so a paragraph that is an
// element of a block can carry anchors below its direct children. These pin that such a paragraph
// renders as it does outside any block: every token in the box substituted once, in each copy of
// the box.
// ReSharper disable PartialTypeWithSinglePart
public partial class TextBoxInBlockTests
{
    public class Customer
    {
        public required string Name { get; init; }

        public bool Active { get; init; }

        public List<string> Tags { get; init; } = [];
    }

    [ParchmentBindable]
    public partial class Model
    {
        public string Title { get; init; } = "Al";

        public bool Show { get; init; }

        public List<Customer> Customers { get; init; } =
        [
            new()
            {
                Name = "c1",
                Active = true,
                Tags = ["one", "two"]
            },
            new()
            {
                Name = "c2",
                Tags = ["three"]
            }
        ];
    }

    public class Line
    {
        [EditableField]
        public string? Name { get; set; }

        public string Code { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class Order
    {
        [EditableField]
        public required List<Line> Lines { get; set; }
    }

    [Test]
    public async Task LoopBodyParagraphAnchoringATextBox()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Host("Dear {{ c.Name }}", TextBox(P("box {{ c.Name }}"))),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("Dear c1 | box c1 | box c1 | Dear c2 | box c2 | box c2");
    }

    // The host has no token, so it has no anchor of its own and only the box's paragraphs do.
    [Test]
    public async Task LoopBodyParagraphWithNoTokenOfItsOwn()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Host("static host", TextBox(P("box {{ c.Name }}"))),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("static host | box c1 | box c1 | static host | box c2 | box c2");
    }

    [Test]
    public async Task IfBranchParagraphAnchoringATextBox()
    {
        using var template = Build(
            P("{% if Show %}"),
            Host("Dear {{ Title }}", TextBox(P("box {{ Title }}"))),
            P("{% endif %}"));

        var texts = await Render(
            template,
            new()
            {
                Show = true
            });

        await Assert.That(texts).IsEqualTo("Dear Al | box Al | box Al");
    }

    [Test]
    public async Task ElsifBranchParagraphAnchoringATextBox()
    {
        using var template = Build(
            P("{% if Show %}"),
            P("shown"),
            P("{% elsif Title == 'Al' %}"),
            Host("host", TextBox(P("box {{ Title }}"))),
            P("{% else %}"),
            P("neither"),
            P("{% endif %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("host | box Al | box Al");
    }

    [Test]
    public async Task ElseBranchParagraphAnchoringATextBox()
    {
        using var template = Build(
            P("{% if Show %}"),
            P("shown"),
            P("{% else %}"),
            Host("host", TextBox(P("box {{ Title }}"))),
            P("{% endif %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("host | box Al | box Al");
    }

    // The tags of the conditional are paragraphs of the box, so each iteration has to find its own
    // copies of them. Left unprocessed they print, and the content prints for every customer.
    [Test]
    public async Task ConditionalInsideATextBoxInALoopBody()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Host("host", TextBox(P("{% if c.Active %}") + P("active {{ c.Name }}") + P("{% endif %}"))),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("host | active c1 | active c1 | host");
    }

    [Test]
    public async Task LoopInsideATextBoxInALoopBody()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Host("{{ c.Name }}", VmlTextBox(P("{% for tag in c.Tags %}") + P("{{ c.Name }} {{ tag }}") + P("{% endfor %}"))),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("c1 | c1 one | c1 two | c2 | c2 three");
    }

    // The inner loop's elements are clones made by the outer one, and the conditional's are the
    // inner loop's: each level indexes what it was handed.
    [Test]
    public async Task TextBoxInAConditionalInANestedLoop()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            P("{% for tag in c.Tags %}"),
            P("{% if c.Active %}"),
            Host("{{ tag }}", TextBox(P("{{ c.Name }} {{ tag }}"))),
            P("{% endif %}"),
            P("{% endfor %}"),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("one | c1 one | c1 one | two | c1 two | c1 two");
    }

    // A table is walked for every paragraph below it, so this never depended on the paragraph case.
    [Test]
    public async Task TableCellParagraphAnchoringATextBoxInALoopBody()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Table(Host("cell {{ c.Name }}", TextBox(P("box {{ c.Name }}")))),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("cell c1 | box c1 | box c1 | cell c2 | box c2 | box c2");
    }

    [Test]
    public async Task TableInsideATextBoxInALoopBody()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Host("host", VmlTextBox(Table(P("cell {{ c.Name }}")))),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("host | cell c1 | host | cell c2");
    }

    // The loop is itself in a text box, and a paragraph of its body anchors another.
    [Test]
    public async Task LoopInsideATextBoxWhoseBodyAnchorsATextBox()
    {
        using var template = Build(
            Host(
                "host",
                VmlTextBox(
                    P("{% for c in Customers %}") +
                    Host("outer {{ c.Name }}", VmlTextBox(P("inner {{ c.Name }}"))) +
                    P("{% endfor %}"))));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("host | outer c1 | inner c1 | outer c2 | inner c2");
    }

    // Only the tag goes from a tag paragraph: a text box anchored in it stays where the tag was,
    // whichever way its block goes, and its tokens are read in the scope the tag paragraph is in.
    // The box is stored twice, as the shape and its fallback.
    [Test]
    public async Task TextBoxAnchoredInATagParagraphStaysWhereTheTagWas()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            Host("{% if c.Active %}", TextBox(P("box {{ c.Name }}"))),
            P("active {{ c.Name }}"),
            P("{% endif %}"),
            P("{% endfor %}"));

        var texts = await Render(template, new());

        await Assert.That(texts).IsEqualTo("box c1 | box c1 | active c1 | box c2 | box c2");
    }

    [Test]
    public async Task HeaderLoopBodyParagraphAnchoringATextBox()
    {
        using var template = BuildWithHeader(
            P("{% for c in Customers %}"),
            Host("Dear {{ c.Name }}", TextBox(P("box {{ c.Name }}"))),
            P("{% endfor %}"));
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var result = await Render(store, new Model());

        await Assert.That(HeaderTexts(result)).IsEqualTo("Dear c1 | box c1 | box c1 | Dear c2 | box c2 | box c2");
    }

    // The markdown flow binds the style source's headers through the same scope tree.
    [Test]
    public async Task MarkdownFlowHeaderBranchParagraphAnchoringATextBox()
    {
        using var styleSource = BuildWithHeader(
            P("{% if Show %}"),
            Host("Dear {{ Title }}", TextBox(P("box {{ Title }}"))),
            P("{% endif %}"));
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# {{ Title }}", styleSource);

        using var result = await Render(
            store,
            new Model
            {
                Show = true
            });

        await Assert.That(HeaderTexts(result)).IsEqualTo("Dear Al | box Al | box Al");
    }

    // An editable collection clones its loop body per item through a path of its own.
    [Test]
    public async Task EditableCollectionItemAnchoringATextBox()
    {
        using var template = Build(
            P("{% for line in Lines %}"),
            Host("{{ line.Name }}", TextBox(P("code {{ line.Code }}"))),
            P("{% endfor %}"));
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Order>(template);

        using var result = await Render(
            store,
            new Order
            {
                Lines =
                [
                    new()
                    {
                        Name = "first",
                        Code = "A1"
                    },
                    new()
                    {
                        Name = "second",
                        Code = "B2"
                    }
                ]
            });

        await Assert.That(Texts(result.MainDocumentPart!.Document!.Body!))
            .IsEqualTo("first | code A1 | code A1 | second | code B2 | code B2");
    }

    // Looking below a paragraph for anchors is paid for per block, and only by a block with a
    // paragraph that has some: not a loop with no text box, not one whose text box is in a table
    // (a table is searched in full anyway), and not a loop that is itself inside a text box.
    [Test]
    public async Task OnlyABlockWhoseParagraphAnchorsATextBoxIsMarked()
    {
        using var template = Build(
            P("{% for c in Customers %}"),
            P("{{ c.Name }}"),
            P("{% endfor %}"),
            P("{% for c in Customers %}"),
            Table(Host("cell", VmlTextBox(P("{{ c.Name }}")))),
            P("{% endfor %}"),
            Host("host", VmlTextBox(P("{% for c in Customers %}") + P("{{ c.Name }}") + P("{% endfor %}"))),
            P("{% if Show %}"),
            Host("host", VmlTextBox(P("{{ Title }}"))),
            P("{% endif %}"),
            P("{% for c in Customers %}"),
            Host("host", VmlTextBox(P("{{ c.Name }}"))),
            P("{% endfor %}"));
        using var document = WordprocessingDocument.Open(template, true);
        var body = document.MainDocumentPart!.Document!.Body!;

        var tree = ScopeTreeBuilder.Build(TokenScanner.Scan(body, "Model", "document"), "Model", "document");

        var marked = tree
            .Select(_ => _ switch
            {
                LoopNode loop => loop.NestedAnchors,
                IfNode conditional => conditional.NestedAnchors,
                _ => throw new InvalidOperationException()
            })
            .ToList();
        await Assert.That(string.Join(" | ", marked)).IsEqualTo("False | False | False | True | True");
    }

    static async Task<string> Render(Stream template, Model model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        using var result = await Render(store, model);
        return Texts(result.MainDocumentPart!.Document!.Body!);
    }

    static async Task<WordprocessingDocument> Render<T>(TemplateStore store, T model)
        where T : notnull
    {
        var output = new MemoryStream();
        await store.Render(model, output);
        output.Position = 0;
        return WordprocessingDocument.Open(output, false);
    }

    static string HeaderTexts(WordprocessingDocument document) =>
        Texts(document.MainDocumentPart!.HeaderParts.Single().Header!);

    static string Texts(OpenXmlElement root) =>
        string.Join(" | ", root.Descendants<Text>().Select(_ => _.Text));

    const string namespaces =
        """xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" """;

    static string P(string text) =>
        Host(text, "");

    // A paragraph with a run of text followed by whatever else is anchored in it.
    static string Host(string text, string anchored) =>
        $"""<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r>{anchored}</w:p>""";

    static string Table(string cellContent) =>
        $"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc>{cellContent}</w:tc></w:tr></w:tbl>";

    // The text box Word writes: DrawingML, with a VML copy of the same content as the fallback.
    static string TextBox(string content) =>
        $"""<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="tb"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback>{VmlShape(content)}</mc:Fallback></mc:AlternateContent></w:r>""";

    static string VmlTextBox(string content) =>
        $"<w:r>{VmlShape(content)}</w:r>";

    static string VmlShape(string content) =>
        $"<w:pict><v:shape><v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape></w:pict>";

    static MemoryStream Build(params string[] blocks)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body($"<w:body {namespaces}>{string.Concat(blocks)}</w:body>"));
        }

        stream.Position = 0;
        return stream;
    }

    static MemoryStream BuildWithHeader(params string[] headerBlocks)
    {
        var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            header.Header = new($"<w:hdr {namespaces}>{string.Concat(headerBlocks)}</w:hdr>");
            main.Document = new(
                new Body(
                    new Paragraph(new Run(new Text("Body"))),
                    new SectionProperties(
                        new HeaderReference
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rIdHeader"
                        })));
        }

        stream.Position = 0;
        return stream;
    }
}
