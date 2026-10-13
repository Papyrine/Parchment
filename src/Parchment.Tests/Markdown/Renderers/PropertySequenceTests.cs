// ReSharper disable PartialTypeWithSinglePart

// The renderers build a fresh w:pPr, w:rPr or w:tcPr through the children constructor, which
// appends in the order given, rather than through the typed setters, which look the position up.
// The order is therefore kept by hand, and these pin it: each shape is checked child by child, and
// the document holding all of them has to validate.
public partial class PropertySequenceTests
{
    [ParchmentBindable]
    public partial class EmptyModel;

    const string everything =
        """
        # Heading

        ## Classed heading {.Title}

        Plain with **bold**, *italic*, ***both***, **_inner italic_**, *__inner bold__*, ~~struck~~,
        ~sub~, ^sup^, ++inserted++, ==marked==, **styled**{.Strong}, ***styled both***{.Strong} and
        `code`, `styled code`{.Mono}.

        A [link](https://example.com/a), a [**bold** link](https://example.com/b), a
        **[link in bold](https://example.com/c)**, a [styled link](https://example.com/d){.Mine}, a
        [`code` link](https://example.com/e) and a *[link in italic](https://example.com/f)*{.Lead}.

        {.Lead}
        Classed paragraph.

        - bullet
        - **bold bullet**

          continuation
          - nested
            - deeper

              deeper continuation
        - ## Heading in an item
        - text

          ### Heading after the text
        -
        - classed item{.Caption}

        1. first
        2. second
           1. inner

        - [ ] open
        - [x] done

        > quote
        >
        > classed quote{.Lead}
        >
        > ## Heading in a quote
        >
        > > nested quote
        > >
        > > ### Heading in a nested quote

        | Left | Centre | Right | None |
        |:-----|:------:|------:|------|
        | a    | b      | c     | d    |
        | **bold** | [link](https://example.com/g) | *it* `code` | ***both*** |
        |      |        |       |      |

        | **Bold** header | [Link](https://example.com/h) header |
        |-----------------|:-------------------------------------|
        | plain           | plain                                |

        +---------+---------+
        | Spanned           |
        +=========+=========+
        | one     | two     |
        +---------+---------+
        | - item  | ## Head |
        +---------+---------+
        """;

    [Test]
    public async Task EveryShapeValidates()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<EmptyModel>(everything);

        using var stream = new MemoryStream();
        await store.Render(new EmptyModel(), stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .Select(_ => $"{_.Path?.XPath}: {_.Description}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    [Test]
    [Arguments("- item", "pStyle numPr contextualSpacing")]
    [Arguments("1. item", "pStyle numPr contextualSpacing")]
    [Arguments("- [x] item", "pStyle numPr contextualSpacing")]
    [Arguments("- item{.Caption}", "pStyle numPr contextualSpacing")]
    [Arguments("-\n- item", "pStyle numPr contextualSpacing")]
    [Arguments("- ## Title", "pStyle numPr")]
    [Arguments("# Title", "pStyle")]
    [Arguments("> text", "pStyle")]
    [Arguments("> ## Title", "pStyle ind")]
    [Arguments("{.Lead}\ntext", "pStyle")]
    public async Task FirstParagraphProperties(string markdown, string expected)
    {
        var paragraph = RendererHarness.RenderMarkdown(markdown).OfType<Paragraph>().First();

        await Assert.That(Sequence(paragraph.ParagraphProperties!)).IsEqualTo(expected);
        await Assert.That(paragraph.FirstChild).IsSameReferenceAs(paragraph.ParagraphProperties);
    }

    [Test]
    [Arguments("- item\n\n  continuation", "pStyle ind contextualSpacing")]
    [Arguments("- item\n\n  ## Title", "pStyle ind")]
    [Arguments("> one\n>\n> > two", "pStyle ind")]
    [Arguments("> one\n>\n> > ## Two", "pStyle ind")]
    [Arguments("- outer\n  - inner", "pStyle numPr contextualSpacing")]
    public async Task SecondParagraphProperties(string markdown, string expected)
    {
        var paragraph = RendererHarness.RenderMarkdown(markdown).OfType<Paragraph>().ElementAt(1);

        await Assert.That(Sequence(paragraph.ParagraphProperties!)).IsEqualTo(expected);
        await Assert.That(paragraph.FirstChild).IsSameReferenceAs(paragraph.ParagraphProperties);
    }

    [Test]
    public async Task PlainParagraphHasNoProperties()
    {
        var paragraph = RendererHarness.RenderMarkdown("text").OfType<Paragraph>().Single();

        await Assert.That(paragraph.ParagraphProperties).IsNull();
    }

    // Each case names the run by its text, so the order of the runs around it does not matter.
    [Test]
    [Arguments("**x**", "b")]
    [Arguments("*x*", "i")]
    [Arguments("~~x~~", "strike")]
    [Arguments("~x~", "vertAlign")]
    [Arguments("^x^", "vertAlign")]
    [Arguments("++x++", "u")]
    [Arguments("==x==", "highlight")]
    [Arguments("***x***", "b i")]
    [Arguments("**_x_**", "b i")]
    [Arguments("*__x__*", "b i")]
    [Arguments("~~**x**~~", "b strike")]
    [Arguments("**~~x~~**", "b strike")]
    [Arguments("==++*x*++==", "i highlight u")]
    [Arguments("**x**{.Strong}", "rStyle b")]
    [Arguments("***x***{.Strong}", "rStyle b i")]
    [Arguments("`x`", "rFonts")]
    [Arguments("`x`{.Mono}", "rStyle rFonts")]
    [Arguments("**`x`**", "rFonts b")]
    [Arguments("**`x`{.Mono}**", "rStyle rFonts b")]
    [Arguments("[x](https://example.com)", "rStyle")]
    [Arguments("[x](https://example.com){.Mine}", "rStyle")]
    [Arguments("[**x**](https://example.com)", "rStyle b")]
    [Arguments("**[x](https://example.com)**", "rStyle b")]
    [Arguments("*[**x**](https://example.com)*", "rStyle b i")]
    [Arguments("[`x`](https://example.com)", "rStyle rFonts")]
    [Arguments("<https://x.example>", "rStyle")]
    [Arguments("<b>**x**</b>", "b")]
    [Arguments("<u>*x*</u>", "i u")]
    public async Task RunProperties(string markdown, string expected)
    {
        var paragraph = RendererHarness.RenderMarkdown(markdown).OfType<Paragraph>().Single();
        var run = paragraph.Descendants<Run>().Single(_ => _.InnerText.Contains('x'));

        await Assert.That(Sequence(run.RunProperties!)).IsEqualTo(expected);
        await Assert.That(run.FirstChild).IsSameReferenceAs(run.RunProperties);
    }

    [Test]
    public async Task EmphasisStyleReplacesTheLinkStyleInPlace()
    {
        var paragraph = RendererHarness.RenderMarkdown("**[x](https://example.com)**{.Strong}").OfType<Paragraph>().Single();
        var properties = paragraph.Descendants<Run>().Single().RunProperties!;

        await Assert.That(Sequence(properties)).IsEqualTo("rStyle b");
        await Assert.That(properties.RunStyle!.Val?.Value).IsEqualTo("Strong");
    }

    [Test]
    public async Task AlignedTableCells()
    {
        const string markdown =
            """
            | Left | **Centre** | Right |
            |:-----|:----------:|------:|
            | a    | *b*        | [c](https://example.com) |
            """;
        var table = RendererHarness.RenderMarkdown(markdown).OfType<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();

        foreach (var cell in table.Descendants<TableCell>())
        {
            // tcPr leads the cell and pPr the paragraph, whichever path built them.
            await Assert.That(cell.FirstChild).IsSameReferenceAs(cell.TableCellProperties);
            await Assert.That(Sequence(cell.TableCellProperties!)).IsEqualTo("tcW");
            var paragraph = cell.Elements<Paragraph>().Single();
            await Assert.That(paragraph.FirstChild).IsSameReferenceAs(paragraph.ParagraphProperties);
            await Assert.That(Sequence(paragraph.ParagraphProperties!)).IsEqualTo("jc");
        }

        var header = rows[0].Descendants<Run>().ToList();
        await Assert.That(header.Select(_ => Sequence(_.RunProperties!))).IsEquivalentTo(["b", "b", "b"]);

        var body = rows[1].Descendants<Run>().ToList();
        await Assert.That(body[0].RunProperties).IsNull();
        await Assert.That(Sequence(body[1].RunProperties!)).IsEqualTo("i");
        await Assert.That(Sequence(body[2].RunProperties!)).IsEqualTo("rStyle");
    }

    // A header cell holding a link or a heading reaches the bold and the centring with properties
    // already in place, so those go in by the schema sequence rather than first or last.
    [Test]
    public async Task HeaderCellFormattingJoinsWhatTheCellAlreadyHolds()
    {
        const string markdown =
            """
            +------------------------------+---------+
            | [link](https://example.com)  | ## Head |
            +==============================+=========+
            | one                          | two     |
            +------------------------------+---------+
            """;
        var table = RendererHarness.RenderMarkdown(markdown).OfType<Table>().Single();
        var cells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();

        await Assert.That(Sequence(cells[0].Descendants<Run>().Single().RunProperties!)).IsEqualTo("rStyle b");
        await Assert.That(Sequence(cells[1].Elements<Paragraph>().Single().ParagraphProperties!)).IsEqualTo("pStyle jc");
    }

    [Test]
    public async Task SpannedCell()
    {
        const string markdown =
            """
            +---------+---------+
            | Spanned           |
            +=========+=========+
            | one     | two     |
            +---------+---------+
            """;
        var table = RendererHarness.RenderMarkdown(markdown).OfType<Table>().Single();
        var cell = table.Descendants<TableCell>().First();

        await Assert.That(cell.FirstChild).IsSameReferenceAs(cell.TableCellProperties);
        await Assert.That(Sequence(cell.TableCellProperties!)).IsEqualTo("gridSpan");
    }

    static string Sequence(OpenXmlElement properties) =>
        string.Join(' ', properties.ChildElements.Select(_ => _.LocalName));
}
