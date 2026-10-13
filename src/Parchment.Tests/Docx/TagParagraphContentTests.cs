using V = DocumentFormat.OpenXml.Vml;

// A block tag paragraph is removed and the host of a solo structural token is replaced, and both
// are recognised by their text alone. These pin that only the tag or the token goes: a section
// break, a page break, a picture or a footnote reference sharing the paragraph stays in the
// document, once and where the author put it.
// ReSharper disable PartialTypeWithSinglePart
public partial class TagParagraphContentTests
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    sealed class HtmlAttribute : Attribute;

    [ParchmentBindable]
    public partial class Model
    {
        public string Name { get; init; } = "Alice";
        public bool Flag { get; init; }
        public IReadOnlyList<Line> Lines { get; init; } = [];
        public string[] Tags { get; init; } = [];

        [Html]
        public string Html { get; init; } = "";

        [ExcelsiorTable]
        public IReadOnlyList<Line>? Rows { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
    }

    [ParchmentBindable]
    public partial class EditableModel
    {
        [Html]
        [EditableField]
        public string Body { get; set; } = "";
    }

    static Model TwoLines =>
        new()
        {
            Lines =
            [
                new()
                {
                    Name = "A"
                },
                new()
                {
                    Name = "B"
                }
            ]
        };

    const int wide = 15840;
    const int narrow = 9000;

    [Test]
    public async Task SectionBreakOnEndForEndsTheSectionAfterTheLastIteration()
    {
        using var template = Build(
            P("first", narrow),
            P("second"),
            P("{% for line in Lines %}"),
            P("{{ line.Name }}"),
            P("{% endfor %}", wide),
            P("third"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo($"[first §{narrow}] [second] [A] [B §{wide}] [third]");
        await AssertValid(doc);
    }

    [Test]
    public async Task SectionBreakOnEndForOfAnEmptyLoopEndsTheSectionBeforeTheLoop()
    {
        using var template = Build(
            P("first", narrow),
            P("second"),
            P("{% for line in Lines %}"),
            P("{{ line.Name }}"),
            P("{% endfor %}", wide),
            P("third"));

        using var doc = await Render(template, new());

        await Assert.That(Outline(doc)).IsEqualTo($"[first §{narrow}] [second §{wide}] [third]");
    }

    // A table cannot carry a section break, so an empty paragraph stays behind it to hold one.
    [Test]
    public async Task SectionBreakOnEndForStaysOnAParagraphAfterATable()
    {
        using var template = Build(
            P("first"),
            P("{% for line in Lines %}"),
            OneCellTable("{{ line.Name }}"),
            P("{% endfor %}", wide),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Lines =
                [
                    new()
                    {
                        Name = "A"
                    }
                ]
            });

        await Assert.That(Outline(doc)).IsEqualTo($"[first] {{table}} [ §{wide}] [second]");
        await AssertValid(doc);
    }

    // The paragraph before already ends a section of its own, so the break keeps its own.
    [Test]
    public async Task SectionBreakStaysWhereThePrecedingParagraphEndsAnotherSection()
    {
        using var template = Build(
            P("first", narrow),
            P("{% for line in Lines %}"),
            P("{{ line.Name }}"),
            P("{% endfor %}", wide),
            P("second"));

        using var doc = await Render(template, new());

        await Assert.That(Outline(doc)).IsEqualTo($"[first §{narrow}] [ §{wide}] [second]");
    }

    [Test]
    public async Task SectionBreakOnForIsWrittenOnceBeforeTheIterations()
    {
        using var template = Build(
            P("first"),
            P("{% for line in Lines %}", wide),
            P("{{ line.Name }}"),
            P("{% endfor %}"),
            P("second"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo($"[first §{wide}] [A] [B] [second]");
    }

    [Test]
    [Arguments(true, "[first] [yes §15840] [second]")]
    [Arguments(false, "[first §15840] [second]")]
    public async Task SectionBreakOnEndIfSurvivesEitherOutcome(bool flag, string expected)
    {
        using var template = Build(
            P("first"),
            P("{% if Flag %}"),
            P("yes"),
            P("{% endif %}", wide),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        await Assert.That(Outline(doc)).IsEqualTo(expected);
    }

    // The else tag belongs to neither branch, so its break is kept whichever one is chosen.
    [Test]
    [Arguments(true, "[first] [yes §15840] [second]")]
    [Arguments(false, "[first §15840] [no] [second]")]
    public async Task SectionBreakOnElseSurvivesEitherOutcome(bool flag, string expected)
    {
        using var template = Build(
            P("first"),
            P("{% if Flag %}"),
            P("yes"),
            P("{% else %}", wide),
            P("no"),
            P("{% endif %}"),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        await Assert.That(Outline(doc)).IsEqualTo(expected);
    }

    // A block inside a loop body is part of what repeats, and so is the break on its tag.
    [Test]
    public async Task SectionBreakOnATagInsideALoopBodyRepeatsWithTheBody()
    {
        using var template = Build(
            P("{% for line in Lines %}"),
            P("{% if line.Name %}"),
            P("{{ line.Name }}"),
            P("{% endif %}", wide),
            P("{% endfor %}"),
            P("last"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo($"[A §{wide}] [B §{wide}] [last]");
    }

    [Test]
    public async Task SectionBreakOnASoloTokenEndsTheSectionAfterWhatTheTokenProduced()
    {
        using var template = Build(
            P("first"),
            P("{{ Html }}", wide),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        await Assert.That(Outline(doc)).IsEqualTo($"[first] [one] [two §{wide}] [second]");
        await AssertValid(doc);
    }

    [Test]
    public async Task SectionBreakOnASoloTokenThatProducedNothingMovesToTheParagraphBefore()
    {
        using var template = Build(
            P("first"),
            P("{{ Rows }}", wide),
            P("second"));

        using var doc = await Render(template, new());

        await Assert.That(Outline(doc)).IsEqualTo($"[first §{wide}] [second]");
    }

    [Test]
    public async Task SectionBreakOnASoloTableTokenStaysOnAParagraphAfterTheTable()
    {
        using var template = Build(
            P("first"),
            P("{{ Rows }}", wide),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Rows =
                [
                    new()
                    {
                        Name = "A"
                    }
                ]
            });

        await Assert.That(Outline(doc)).IsEqualTo($"[first] {{table}} [ §{wide}] [second]");
        await AssertValid(doc);
    }

    // The last list item takes the break beside its numbering, in the order the schema wants.
    [Test]
    public async Task SectionBreakOnASoloListTokenLandsOnTheLastItem()
    {
        using var template = Build(
            P("{{ Tags }}", wide),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Tags = ["one", "two"]
            });

        await Assert.That(Outline(doc)).IsEqualTo($"[one] [two §{wide}] [second]");
        var last = BodyOf(doc).Elements<Paragraph>().ElementAt(1);
        await Assert.That(last.ParagraphProperties!.NumberingProperties).IsNotNull();
        await AssertValid(doc);
    }

    [Test]
    [Arguments("{% for line in Lines %}", "[<page>] [A] [B] [end]")]
    [Arguments("{% endfor %}", "[A] [B] [<page>] [end]")]
    public async Task PageBreakOnALoopTagIsKeptOnceWhereTheTagWas(string tag, string expected)
    {
        using var template = Build(
            Tag("{% for line in Lines %}", tag),
            P("{{ line.Name }}"),
            Tag("{% endfor %}", tag),
            P("end"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo(expected);
        await AssertValid(doc);
    }

    [Test]
    [Arguments("{% if Flag %}", true, "[<page>] [yes] [end]")]
    [Arguments("{% if Flag %}", false, "[<page>] [no] [end]")]
    [Arguments("{% else %}", true, "[yes] [<page>] [end]")]
    [Arguments("{% else %}", false, "[<page>] [no] [end]")]
    [Arguments("{% endif %}", true, "[yes] [<page>] [end]")]
    [Arguments("{% endif %}", false, "[no] [<page>] [end]")]
    public async Task PageBreakOnAConditionalTagIsKeptWhicheverBranchIsChosen(string tag, bool flag, string expected)
    {
        using var template = Build(
            Tag("{% if Flag %}", tag),
            P("yes"),
            Tag("{% else %}", tag),
            P("no"),
            Tag("{% endif %}", tag),
            P("end"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        await Assert.That(Outline(doc)).IsEqualTo(expected);
    }

    [Test]
    [Arguments(true, "[yes] [<page>]")]
    [Arguments(false, "[<page>]")]
    public async Task PageBreakOnAnUnchosenElsifIsKept(bool flag, string expected)
    {
        using var template = Build(
            P("{% if Flag %}"),
            P("yes"),
            new Paragraph(PageBreakRun(), TextRun("{% elsif Flag %}")),
            P("other"),
            P("{% endif %}"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        await Assert.That(Outline(doc)).IsEqualTo(expected);
    }

    [Test]
    public async Task PictureOnATagParagraphIsKeptAndNotRepeated()
    {
        using var template = Build(
            new Paragraph(TextRun("{% for line in Lines %}"), PictureRun()),
            P("{{ line.Name }}"),
            P("{% endfor %}"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo("[<picture>] [A] [B]");
        await AssertValid(doc);
    }

    [Test]
    public async Task FootnoteReferenceOnATagParagraphIsKept()
    {
        using var template = Build(
            P("{% if Flag %}"),
            P("yes"),
            new Paragraph(TextRun("{% endif %}"), NoteReferenceRun()));
        AddFootnote(template, P("the note"));

        using var doc = await Render(template, new());

        await Assert.That(Outline(doc)).IsEqualTo("[<note>]");
        await AssertValid(doc);
    }

    // A tag paragraph that stays for what it holds keeps the section break it carries too.
    [Test]
    public async Task TagParagraphKeptForItsContentKeepsItsSectionBreak()
    {
        var close = new Paragraph(
            new ParagraphProperties
            {
                SectionProperties = Section(wide)
            },
            PageBreakRun(),
            TextRun("{% endfor %}"));
        using var template = Build(
            P("{% for line in Lines %}"),
            P("{{ line.Name }}"),
            close,
            P("end"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo($"[A] [B] [<page> §{wide}] [end]");
    }

    // What goes with a tag paragraph as it always has: white space, the page boundary Word
    // cached, a bookmark, a field, a comment on the tag, formatting that is not a page break.
    [Test]
    public async Task TagParagraphHoldingNothingToKeepIsStillRemoved()
    {
        var open = new Paragraph(
            new BookmarkStart
            {
                Id = "900",
                Name = "_GoBack"
            },
            new BookmarkEnd
            {
                Id = "900"
            },
            new Run(new LastRenderedPageBreak(), new TabChar(), new Text("{% for line in Lines %}"), new Break()),
            FieldMark(FieldCharValues.Begin),
            new Run(new FieldCode(" XE \"entry\" ")),
            FieldMark(FieldCharValues.End));
        var close = new Paragraph(
            new ParagraphProperties(
                new KeepNext(),
                new PageBreakBefore
                {
                    Val = false
                }),
            new CommentRangeStart
            {
                Id = "5"
            },
            TextRun("{% endfor %}"),
            new CommentRangeEnd
            {
                Id = "5"
            },
            new Run(
                new CommentReference
                {
                    Id = "5"
                }));
        using var template = Build(open, P("{{ line.Name }}"), close);

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo("[A] [B]");
    }

    // A token typed over the result of a merge field is the usual way a field ends up around
    // one, so the field goes with the token whether or not something else keeps the paragraph.
    [Test]
    [Arguments(false, "[one] [two]")]
    [Arguments(true, "[<page>] [one] [two]")]
    public async Task FieldAroundASoloTokenGoesWithTheToken(bool pageBreak, string expected)
    {
        var host = new Paragraph(
            FieldMark(FieldCharValues.Begin),
            new Run(new FieldCode(" MERGEFIELD Html ")),
            FieldMark(FieldCharValues.Separate),
            TextRun("{{ Html }}"),
            FieldMark(FieldCharValues.End));
        if (pageBreak)
        {
            host.PrependChild(PageBreakRun());
        }

        using var template = Build(host);

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        await Assert.That(Outline(doc)).IsEqualTo(expected);
        await Assert.That(BodyOf(doc).Descendants<FieldChar>().Count()).IsEqualTo(0);
        await Assert.That(BodyOf(doc).Descendants<FieldCode>().Count()).IsEqualTo(0);
    }

    // "Page break before" is the same break written as paragraph formatting.
    [Test]
    public async Task PageBreakBeforeFormattingOnATagParagraphIsKept()
    {
        using var template = Build(
            P("{% for line in Lines %}"),
            P("{{ line.Name }}"),
            new Paragraph(new ParagraphProperties(new PageBreakBefore()), TextRun("{% endfor %}")),
            P("end"));

        using var doc = await Render(template, TwoLines);

        await Assert.That(Outline(doc)).IsEqualTo("[A] [B] [<page before>] [end]");
    }

    // The break is ahead of everything the host is replaced by, and is not repeated behind it.
    [Test]
    public async Task PageBreakBeforeFormattingOnASoloTokenStaysAheadOfWhatItProduced()
    {
        using var template = Build(
            new Paragraph(
                new ParagraphProperties(new PageBreakBefore()),
                TextRun("{{ Html }}"),
                NoteReferenceRun()));
        AddFootnote(template, P("the note"));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        await Assert.That(Outline(doc)).IsEqualTo("[<page before>] [one] [two] [<note>]");
        await AssertValid(doc);
    }

    [Test]
    public async Task PageBreakBeforeASoloTokenStaysAheadOfWhatItProduced()
    {
        using var template = Build(
            P("first"),
            new Paragraph(PageBreakRun(), TextRun("{{ Html }}")),
            P("second"));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        await Assert.That(Outline(doc)).IsEqualTo("[first] [<page>] [one] [two] [second]");
        await AssertValid(doc);
    }

    [Test]
    public async Task FootnoteReferenceAfterASoloTokenStaysBehindWhatItProduced()
    {
        using var template = Build(new Paragraph(TextRun("{{ Html }}"), NoteReferenceRun()));
        AddFootnote(template, P("the note"));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        await Assert.That(Outline(doc)).IsEqualTo("[one] [two] [<note>]");
        await AssertValid(doc);
    }

    [Test]
    public async Task ContentOnBothSidesOfASoloTokenStaysOnItsSide()
    {
        using var template = Build(new Paragraph(PictureRun(), TextRun("{{ Tags }}"), PageBreakRun()));

        using var doc = await Render(
            template,
            new()
            {
                Tags = ["one", "two"]
            });

        await Assert.That(Outline(doc)).IsEqualTo("[<picture>] [one] [two] [<page>]");
        // The items are still list items: what shared the host is beside them, not in them.
        await Assert.That(BodyOf(doc).Descendants<NumberingProperties>().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task PageBreakBesideASoloTableTokenIsKept()
    {
        using var template = Build(new Paragraph(PageBreakRun(), TextRun("{{ Rows }}")));

        using var doc = await Render(
            template,
            new()
            {
                Rows =
                [
                    new()
                    {
                        Name = "A"
                    }
                ]
            });

        await Assert.That(Outline(doc)).IsEqualTo("[<page>] {table}");
    }

    [Test]
    public async Task ContentBesideASoloTokenThatProducedNothingIsKept()
    {
        using var template = Build(
            new Paragraph(PageBreakRun(), TextRun("{{ Rows }}")),
            P("after"));

        using var doc = await Render(template, new());

        await Assert.That(Outline(doc)).IsEqualTo("[<page>] [after]");
    }

    // An editable rich-text field replaces its host with a block-level control the same way.
    [Test]
    public async Task PageBreakBesideAnEditableHtmlTokenIsKept()
    {
        using var template = Build(new Paragraph(PageBreakRun(), TextRun("{{ Body }}")));
        var store = new TemplateStore();
        store.RegisterDocxTemplate<EditableModel>(template);

        var stream = new MemoryStream();
        await store.Render(
            new EditableModel
            {
                Body = "<p>one</p>"
            },
            stream);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);

        await Assert.That(Outline(doc)).IsEqualTo("[<page>]");
        await Assert.That(BodyOf(doc).Elements<SdtBlock>().Single().InnerText).IsEqualTo("one");
        await AssertValid(doc);
    }

    // A text box anchored on a tag paragraph stays with it, outside the block, and its own
    // tokens are still substituted.
    [Test]
    [Arguments(true, "[<box Alice>] [yes]")]
    [Arguments(false, "[<box Alice>]")]
    public async Task TextBoxOnATagParagraphIsKeptAndSubstituted(bool flag, string expected)
    {
        using var template = Build(
            new Paragraph(TextRun("{% if Flag %}"), new Run(TextBoxShape("{{ Name }}"))),
            P("yes"),
            P("{% endif %}"));

        using var doc = await Render(
            template,
            new()
            {
                Flag = flag
            });

        await Assert.That(Outline(doc)).IsEqualTo(expected);
    }

    [Test]
    public async Task TextBoxOnASoloTokenHostIsKeptAndSubstituted()
    {
        using var template = Build(new Paragraph(TextRun("{{ Html }}"), new Run(TextBoxShape("{{ Name }}"))));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        await Assert.That(Outline(doc)).IsEqualTo("[one] [two] [<box Alice>]");
    }

    [Test]
    public async Task PageBreakOnATagParagraphInATableCellIsKept()
    {
        using var template = Build(
            new Table(
                new TableProperties(
                    new TableWidth
                    {
                        Type = TableWidthUnitValues.Auto
                    }),
                new TableGrid(new GridColumn()),
                new TableRow(
                    new TableCell(
                        new Paragraph(PageBreakRun(), TextRun("{% if Flag %}")),
                        P("yes"),
                        P("{% endif %}")))));

        using var doc = await Render(template, new());

        var cell = BodyOf(doc).Descendants<TableCell>().Single();
        await Assert.That(Outline(cell)).IsEqualTo("[<page>]");
    }

    // A footnote's first paragraph opens with the mark that numbers it.
    [Test]
    public async Task FootnoteMarkBesideASoloTokenInAFootnoteIsKept()
    {
        using var template = Build(new Paragraph(TextRun("body"), NoteReferenceRun()));
        AddFootnote(template, new(new Run(new FootnoteReferenceMark()), TextRun("{{ Html }}")));

        using var doc = await Render(
            template,
            new()
            {
                Html = "<p>one</p><p>two</p>"
            });

        var note = doc.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().Single();
        await Assert.That(Outline(note)).IsEqualTo("[<mark>] [one] [two]");
    }

    [Test]
    public async Task PictureOnATagParagraphInAHeaderIsKept()
    {
        using var template = Build(P("body"));
        AddHeader(
            template,
            new Paragraph(PictureRun(), TextRun("{% if Flag %}")),
            P("for {{ Name }}"),
            P("{% endif %}"));

        using var doc = await Render(template, new());

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await Assert.That(Outline(header)).IsEqualTo("[<picture>]");
    }

    // A markdown template runs the tokens of its style source's headers through the same runner.
    [Test]
    public async Task PictureOnATagParagraphInAMarkdownStyleSourceHeaderIsKept()
    {
        using var styleSource = Build();
        AddHeader(
            styleSource,
            new Paragraph(PictureRun(), TextRun("{% if Flag %}")),
            P("for {{ Name }}"),
            P("{% endif %}"));
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("# {{ Name }}", styleSource);

        var stream = new MemoryStream();
        await store.Render(new Model(), stream);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);

        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;
        await Assert.That(Outline(header)).IsEqualTo("[<picture>]");
    }

    static async Task<WordprocessingDocument> Render(Stream template, Model model)
    {
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }

    static async Task AssertValid(WordprocessingDocument doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(doc)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}")
            .ToList();
        await Assert.That(errors).IsEmpty();
    }

    static Body BodyOf(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!;

    static string Outline(WordprocessingDocument doc) =>
        Outline(BodyOf(doc));

    /// <summary>
    /// The blocks of <paramref name="container"/> in order: <c>[text]</c> for a paragraph, with a
    /// marker for each thing it holds besides text and <c>§width</c> when it ends a section, and
    /// <c>{table}</c> for a table.
    /// </summary>
    static string Outline(OpenXmlElement container)
    {
        var blocks = new List<string>();
        foreach (var block in container.ChildElements)
        {
            if (block is Table)
            {
                blocks.Add("{table}");
                continue;
            }

            if (block is not Paragraph paragraph)
            {
                continue;
            }

            var builder = new StringBuilder();
            if (paragraph.ParagraphProperties?.PageBreakBefore != null)
            {
                builder.Append("<page before>");
            }

            Describe(paragraph, builder);
            if (paragraph.ParagraphProperties?.SectionProperties is { } section)
            {
                builder.Append($" §{section.GetFirstChild<PageSize>()!.Width!.Value}");
            }

            blocks.Add($"[{builder}]");
        }

        return string.Join(' ', blocks);
    }

    static void Describe(OpenXmlElement element, StringBuilder builder)
    {
        foreach (var child in element.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    builder.Append(text.Text);
                    break;
                case Break @break when @break.Type?.Value == BreakValues.Page:
                    builder.Append("<page>");
                    break;
                case Drawing:
                    builder.Append("<picture>");
                    break;
                case Picture box:
                    builder.Append($"<box {box.InnerText}>");
                    break;
                case FootnoteReference:
                    builder.Append("<note>");
                    break;
                case FootnoteReferenceMark:
                    builder.Append("<mark>");
                    break;
                case ParagraphProperties:
                    break;
                default:
                    Describe(child, builder);
                    break;
            }
        }
    }

    /// <summary>
    /// The shared fixture with <paramref name="blocks"/> as its body, ahead of the document's own
    /// section properties.
    /// </summary>
    static MemoryStream Build(params OpenXmlElement[] blocks)
    {
        var stream = DocxTemplateBuilder.Build();
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var body = BodyOf(doc);
            var last = body.GetFirstChild<SectionProperties>()!;
            foreach (var block in blocks)
            {
                body.InsertBefore(block, last);
            }
        }

        stream.Position = 0;
        return stream;
    }

    static void AddFootnote(MemoryStream template, Paragraph paragraph)
    {
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var part = doc.MainDocumentPart!.AddNewPart<FootnotesPart>("rIdFootnotes");
            part.Footnotes = new(
                new Footnote(paragraph)
                {
                    Id = 1
                });
        }

        template.Position = 0;
    }

    static void AddHeader(MemoryStream template, params OpenXmlElement[] blocks)
    {
        using (var doc = WordprocessingDocument.Open(template, true))
        {
            var main = doc.MainDocumentPart!;
            var part = main.AddNewPart<HeaderPart>("rIdHeader");
            part.Header = new(blocks);
            BodyOf(doc).GetFirstChild<SectionProperties>()!.PrependChild(
                new HeaderReference
                {
                    Type = HeaderFooterValues.Default,
                    Id = "rIdHeader"
                });
        }

        template.Position = 0;
    }

    static Paragraph P(string text) =>
        new(TextRun(text));

    /// <summary>
    /// A paragraph that ends a section whose page is <paramref name="width"/> wide.
    /// </summary>
    static Paragraph P(string text, int width) =>
        new(
            new ParagraphProperties
            {
                SectionProperties = Section(width)
            },
            TextRun(text));

    /// <summary>
    /// A tag paragraph, with a page break ahead of the tag when it is the one under test.
    /// </summary>
    static Paragraph Tag(string tag, string withPageBreak)
    {
        if (tag == withPageBreak)
        {
            return new(PageBreakRun(), TextRun(tag));
        }

        return P(tag);
    }

    static SectionProperties Section(int width) =>
        new(
            new PageSize
            {
                Width = (uint)width,
                Height = 12240
            });

    static Run TextRun(string text) =>
        new(
            new Text(text)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

    static Run PageBreakRun() =>
        new(
            new Break
            {
                Type = BreakValues.Page
            });

    static Run FieldMark(FieldCharValues type) =>
        new(
            new FieldChar
            {
                FieldCharType = type
            });

    static Run NoteReferenceRun() =>
        new(
            new FootnoteReference
            {
                Id = 1
            });

    static Run PictureRun() =>
        new(
            """
            <w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:inline></w:drawing></w:r>
            """);

    static Picture TextBoxShape(string text) =>
        new(
            new V.Shape(
                new V.TextBox(
                    new TextBoxContent(P(text)))));

    static Table OneCellTable(string text) =>
        new(
            new TableProperties(
                new TableWidth
                {
                    Type = TableWidthUnitValues.Auto
                }),
            new TableGrid(new GridColumn()),
            new TableRow(new TableCell(P(text))));
}
