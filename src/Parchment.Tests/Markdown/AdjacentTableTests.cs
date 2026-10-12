using DocumentFormat.OpenXml.Validation;

// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// Word reads two <c>w:tbl</c> siblings with nothing between them as one table, so wherever
/// rendered markdown would leave two tables side by side an empty paragraph has to keep them apart.
/// </summary>
public partial class AdjacentTableTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public IReadOnlyList<Row> Rows { get; init; } = [];

        public string Html { get; init; } = "";

        public IReadOnlyList<TokenValue> Values { get; init; } = [];

        [ExcelsiorTable]
        public IReadOnlyList<Line>? Lines { get; init; }

        [ExcelsiorTable]
        public IReadOnlyList<Line>? Missing { get; init; }

        [StringSyntax("markdown")]
        public string Notes { get; init; } = "";
    }

    public class Row
    {
        public required string Name { get; init; }
    }

    public class Line
    {
        public required string Description { get; init; }
    }

    const string pipeTable =
        """
        | a | b |
        |---|---|
        | 1 | 2 |
        """;

    const string htmlTable = "<table><tr><td>h</td></tr></table>";

    static IReadOnlyList<Line> Lines() =>
    [
        new()
        {
            Description = "Site survey"
        }
    ];

    static Table BuildTable() =>
        new(
            new TableProperties(
                new TableWidth
                {
                    Width = "0",
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(
                new GridColumn()),
            new TableRow(
                new TableCell(
                    new Paragraph(
                        new Run(
                            new Text("raw"))))));

    static async Task<MemoryStream> Render(string markdown, Model? model = null)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>(markdown, styleSource);

        var stream = new MemoryStream();
        await store.Render(model ?? new Model(), stream);
        stream.Position = 0;
        return stream;
    }

    static async Task<MemoryStream> RenderDocx(string content, Model model)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }

    static async Task<string> RenderBlocks(string markdown, Model? model = null)
    {
        using var stream = await Render(markdown, model);
        using var doc = WordprocessingDocument.Open(stream, false);
        await AssertValid(doc);
        return Blocks(doc.MainDocumentPart!.Document!.Body!);
    }

    // The blocks a container holds, in order: "tbl" for a table, "p" for an empty paragraph and
    // "text" for one that says something, so a separator cannot be mistaken for content.
    static string Blocks(OpenXmlElement container) =>
        string.Join(
            ' ',
            container.ChildElements
                .Where(_ => _ is not (SectionProperties or TableCellProperties))
                .Select(Describe));

    static string Describe(OpenXmlElement element)
    {
        if (element is Paragraph { InnerText.Length: > 0 })
        {
            return "text";
        }

        return element.LocalName;
    }

    static async Task AssertValid(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task PipeTablesSeparatedByABlankLineStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            | a | b |
            |---|---|
            | 1 | 2 |

            | c | d | e |
            |---|---|---|
            | 3 | 4 | 5 |
            """);

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    [Test]
    public async Task TablesWithAParagraphBetweenThemGetNoSeparator()
    {
        var blocks = await RenderBlocks(
            $"""
             before

             {pipeTable}

             between

             {pipeTable}

             after
             """);

        await Assert.That(blocks).IsEqualTo("text tbl text tbl text");
    }

    // Liquid has rendered by the time markdown is parsed, so each iteration is one more table in
    // the source and nothing marks where one iteration ended.
    [Test]
    public async Task TablesProducedByALoopStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            {% for row in Rows %}
            | {{ row.Name }} | value |
            |---|---|
            | 1 | 2 |

            {% endfor %}
            """,
            new()
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

        await Assert.That(blocks).IsEqualTo("tbl p tbl p tbl");
    }

    [Test]
    public async Task HtmlTableBlockStaysSeparateFromPipeTables()
    {
        var blocks = await RenderBlocks(
            $"""
             {pipeTable}

             {htmlTable}

             {pipeTable}
             """);

        await Assert.That(blocks).IsEqualTo("tbl p tbl p tbl");
    }

    // A comment block renders nothing, so it is no more between the tables than the blank lines are.
    [Test]
    public async Task HtmlCommentBetweenTablesLeavesThemSeparate()
    {
        var blocks = await RenderBlocks(
            $"""
             {pipeTable}

             <!-- a note to the author -->

             {pipeTable}
             """);

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    [Test]
    public async Task TablesInOneHtmlBlockStaySeparate()
    {
        var blocks = await RenderBlocks(htmlTable + htmlTable);

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    [Test]
    public async Task TablesInOneHtmlValueStaySeparate()
    {
        var blocks = await RenderBlocks(
            "{{ Html | html }}",
            new()
            {
                Html = htmlTable + htmlTable
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    // The value is parked and its marker's paragraph swapped for the table once markdown has been
    // parsed, so the tables only meet when the marker's paragraph goes.
    [Test]
    public async Task HtmlValueTableStaysSeparateFromPipeTables()
    {
        var blocks = await RenderBlocks(
            $$$"""
               {{{pipeTable}}}

               {{ Html | html }}

               {{{pipeTable}}}
               """,
            new()
            {
                Html = htmlTable
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl p tbl");
    }

    [Test]
    public async Task HtmlTokenTablesProducedByALoopStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            {% for value in Values %}
            {{ value }}

            {% endfor %}
            """,
            new()
            {
                Values =
                [
                    new HtmlToken(htmlTable),
                    new HtmlToken(htmlTable),
                    new HtmlToken(htmlTable)
                ]
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl p tbl");
    }

    [Test]
    public async Task OpenXmlTokenTableStaysSeparateFromPipeTables()
    {
        var blocks = await RenderBlocks(
            $$$"""
               {{{pipeTable}}}

               {{ Values[0] }}

               {{{pipeTable}}}
               """,
            new()
            {
                Values =
                [
                    new OpenXmlToken(_ => [BuildTable(), BuildTable()])
                ]
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl p tbl p tbl");
    }

    // Nothing takes the marker's place, so its paragraph was all that kept the two tables apart.
    [Test]
    public async Task TokenRenderingNothingBetweenTablesLeavesThemSeparate()
    {
        var blocks = await RenderBlocks(
            $$$"""
               {{{pipeTable}}}

               {{ Values[0] }}

               {{{pipeTable}}}
               """,
            new()
            {
                Values =
                [
                    new OpenXmlToken(_ => [])
                ]
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    // An Excelsior table is followed by its emptied host paragraph, so only the table before it
    // can touch it. Word happens to keep two adjacent tables apart when their table styles differ,
    // which is all that separated an unstyled pipe table from Excelsior's TableGrid; a pipe table
    // carrying that style had nothing.
    [Test]
    [Arguments("")]
    [Arguments("{.TableGrid}\n")]
    public async Task PipeTableThenExcelsiorTableStaySeparate(string style)
    {
        var blocks = await RenderBlocks(
            $$$"""
               {{{style}}}{{{pipeTable}}}

               {{ Lines }}
               """,
            new()
            {
                Lines = Lines()
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl p");
    }

    [Test]
    public async Task ExcelsiorTableThenPipeTableKeepTheOneParagraphBetweenThem()
    {
        var blocks = await RenderBlocks(
            $$$"""
               {{ Lines }}

               {{{pipeTable}}}
               """,
            new()
            {
                Lines = Lines()
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    [Test]
    public async Task HtmlValueTableThenExcelsiorTableStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            {{ Html | html }}

            {{ Lines }}
            """,
            new()
            {
                Html = htmlTable,
                Lines = Lines()
            });

        await Assert.That(blocks).IsEqualTo("tbl p tbl p");
    }

    // A null collection renders nothing, and between two tables its marker's paragraph was the
    // only thing separating them.
    [Test]
    public async Task NullExcelsiorTableBetweenTablesLeavesThemSeparate()
    {
        var blocks = await RenderBlocks(
            $$$"""
               {{{pipeTable}}}

               {{ Missing }}

               {{{pipeTable}}}
               """);

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    // A list item's blocks are written into the container around the list, so the tables of one
    // item are siblings there, and so are the tables of two neighbouring items.
    [Test]
    public async Task TablesInAListItemStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            - item

              | a | b |
              |---|---|
              | 1 | 2 |

              | c | d |
              |---|---|
              | 3 | 4 |
            """);

        await Assert.That(blocks).IsEqualTo("text tbl p tbl");
    }

    // An item that opens with a table gets an empty paragraph ahead of it to carry its marker, so
    // the second item's marker paragraph is what sits between the two tables.
    [Test]
    public async Task TablesOpeningNeighbouringListItemsStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            - | a | b |
              |---|---|
              | 1 | 2 |
            - | c | d |
              |---|---|
              | 3 | 4 |
            """);

        await Assert.That(blocks).IsEqualTo("p tbl p tbl");
    }

    [Test]
    public async Task TablesInABlockQuoteStaySeparate()
    {
        var blocks = await RenderBlocks(
            """
            > | a | b |
            > |---|---|
            > | 1 | 2 |
            >
            > | c | d |
            > |---|---|
            > | 3 | 4 |
            """);

        await Assert.That(blocks).IsEqualTo("tbl p tbl");
    }

    const string gridTableHoldingTwoTables =
        """
        +---------------+---------------+
        | Head          | Other         |
        +===============+===============+
        | | i | j |     | plain         |
        | |---|---|     |               |
        | | 1 | 2 |     |               |
        |               |               |
        | | k | l |     |               |
        | |---|---|     |               |
        | | 3 | 4 |     |               |
        +---------------+---------------+
        """;

    // The cell that holds the nested tables: the first cell of the grid table's second row.
    static TableCell NestingCell(Body body) =>
        body.Elements<Table>()
            .Single()
            .Elements<TableRow>()
            .Last()
            .Elements<TableCell>()
            .First();

    // The trailing paragraph is the one Word wants closing a cell that would otherwise end with a
    // table.
    [Test]
    public async Task TablesInAGridTableCellStaySeparate()
    {
        using var stream = await Render(gridTableHoldingTwoTables);
        using var doc = WordprocessingDocument.Open(stream, false);
        await AssertValid(doc);

        var cell = NestingCell(doc.MainDocumentPart!.Document!.Body!);

        await Assert.That(Blocks(cell)).IsEqualTo("tbl p tbl p");
    }

    // A grid table is laid out by column position, and it is the marker standing in for the value
    // that markdown sees: two characters longer than the token, which is why the token's row
    // closes two columns short of the others.
    [Test]
    public async Task HtmlValueTableInAGridTableCellStaysSeparateFromTheTableBeforeIt()
    {
        using var stream = await Render(
            """
            +-----------------------+-------+
            | Head                  | Other |
            +=======================+=======+
            | | i | j |             | plain |
            | |---|---|             |       |
            | | 1 | 2 |             |       |
            |                       |       |
            | {{ Html | html }}   |       |
            +-----------------------+-------+
            """,
            new()
            {
                Html = htmlTable
            });
        using var doc = WordprocessingDocument.Open(stream, false);
        await AssertValid(doc);

        var cell = NestingCell(doc.MainDocumentPart!.Document!.Body!);

        await Assert.That(Blocks(cell)).IsEqualTo("tbl p tbl p");
    }

    // The docx flow renders a markdown value through the same renderer and splices what comes back
    // into the template, so the separators have to arrive with it and not be added a second time.
    [Test]
    public async Task MarkdownValueHoldingAdjacentTablesKeepsThemSeparateInDocxFlow()
    {
        using var stream = await RenderDocx(
            """
            Before

            {{ Notes }}

            After
            """,
            new()
            {
                Notes = pipeTable + "\n\n" + pipeTable
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Blocks(doc.MainDocumentPart!.Document!.Body!)).IsEqualTo("text tbl p tbl text");
    }

    [Test]
    public async Task MarkdownValueHoldingAdjacentTablesInACellKeepsThemSeparateInDocxFlow()
    {
        using var stream = await RenderDocx(
            "{{ Notes }}",
            new()
            {
                Notes = gridTableHoldingTwoTables
            });
        using var doc = WordprocessingDocument.Open(stream, false);

        var cell = NestingCell(doc.MainDocumentPart!.Document!.Body!);

        await Assert.That(Blocks(cell)).IsEqualTo("tbl p tbl p");
    }
}
