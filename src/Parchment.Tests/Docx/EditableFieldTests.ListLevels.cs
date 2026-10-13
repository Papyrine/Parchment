using System.Xml.Linq;

public partial class EditableFieldTests
{
    // One list paragraph as Word stores it: its w:ilvl and whether its numbering is ordered.
    // A null entry in a sequence is a plain paragraph between list items.
    record ListLine(int Level, bool Ordered = false);

    [Test]
    public async Task HtmlEditableListLevelJumpUpAndBackExtractsNested()
    {
        var extracted = await ExtractListLines(new(0), new(2), new(0));

        await Assert.That(extracted)
            .IsEqualTo("<ul><li>i0<ul><li><ul><li>i1</li></ul></li></ul></li><li>i2</li></ul>");
    }

    [Test]
    public async Task HtmlEditableListStartingAboveLevelZeroExtractsNested()
    {
        var extracted = await ExtractListLines(new(1), new(0));

        await Assert.That(extracted)
            .IsEqualTo("<ul><li><ul><li>i0</li></ul></li><li>i1</li></ul>");
    }

    [Test]
    public async Task HtmlEditableListLevelJumpDownExtractsNested()
    {
        var extracted = await ExtractListLines(new(0), new(1), new(2), new(0));

        await Assert.That(extracted)
            .IsEqualTo("<ul><li>i0<ul><li>i1<ul><li>i2</li></ul></li></ul></li><li>i3</li></ul>");
    }

    [Test]
    public async Task HtmlEditableListKindChangeAfterLevelJumpExtractsNested()
    {
        var extracted = await ExtractListLines(new(0), new(2), new(1, true), new(0, true));

        await Assert.That(extracted)
            .IsEqualTo("<ul><li>i0<ul><li><ul><li>i1</li></ul></li></ul><ol><li>i2</li></ol></li></ul><ol><li>i3</li></ol>");
    }

    [Test]
    public async Task HtmlEditableDeepListAtEndAfterParagraphExtractsNested()
    {
        var extracted = await ExtractListLines(new(2), null, new(1, true), new(3));

        await Assert.That(extracted)
            .IsEqualTo(
                "<ul><li><ul><li><ul><li>i0</li></ul></li></ul></li></ul>" +
                "<p>i1</p>" +
                "<ol><li><ol><li>i2<ul><li><ul><li>i3</li></ul></li></ul></li></ol></li></ol>");
    }

    [Test]
    public async Task HtmlEditableListLevelSequencesExtractWellFormed()
    {
        using var rendered = await RenderModel(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = listSeed
            });

        foreach (var lines in ListLineSequences())
        {
            var html = ExtractListLinesFrom(rendered, lines);
            var label = string.Join(",", lines.Select(DescribeLine));

            // Throws on anything that is not well-formed.
            var root = XElement.Parse($"<root>{html}</root>");
            foreach (var element in root.Descendants())
            {
                var name = element.Name.LocalName;
                var parent = element.Parent!.Name.LocalName;
                if (name == "li")
                {
                    await Assert.That(parent is "ul" or "ol").IsTrue().Because(label);
                    continue;
                }

                if (name is "ul" or "ol")
                {
                    await Assert.That(parent is "li" or "root").IsTrue().Because(label);
                    await Assert.That(element.Nodes().All(_ => _ is XElement { Name.LocalName: "li" })).IsTrue().Because(label);
                }
            }

            // Every item keeps its text, order, depth and kind.
            await Assert.That(DescribeHtml(root)).IsEqualTo(label);
        }
    }

    [Test]
    public async Task HtmlEditableListLevelJumpsKeepTheirLevelsWhenRenderedAgain()
    {
        ListLine?[][] sequences =
        [
            [new(0), new(2), new(0)],
            [new(1), new(0)],
            [new(2), new(1, true), new(3), new(0, true)],
            [new(0), new(3), null, new(2, true)],
            [new(0, true), new(1), new(1, true), new(3, true)]
        ];

        foreach (var lines in sequences)
        {
            var html = await ExtractListLines(lines);
            var label = string.Join(",", lines.Select(DescribeLine));

            using var again = await RenderModel(
                "{{ Body }}",
                new EditableArticle
                {
                    Title = "T",
                    Body = html!
                });

            // The paragraphs carrying text sit at the levels the user gave them.
            using (var doc = WordprocessingDocument.Open(again, false))
            {
                var levels = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body")
                    .Descendants<Paragraph>()
                    .Where(_ => _.InnerText.Length > 0)
                    .Select(_ => _.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value)
                    .ToList();
                var expected = lines
                    .Select(_ => _?.Level)
                    .ToList();
                await Assert.That(string.Join(",", levels)).IsEqualTo(string.Join(",", expected)).Because(label);
            }

            // And the extracted html is a fixed point.
            again.Position = 0;
            var model = new EditableArticle
            {
                Title = "T",
                Body = ""
            };
            ParchmentExtractor.Extract<EditableArticle>(again).ApplyTo(model);
            await Assert.That(model.Body).IsEqualTo(html).Because(label);
        }
    }

    const string listSeed = "<ul><li>a</li></ul><ol><li>b</li></ol>";

    // Every sequence of one to four list paragraphs over levels 0-3, in three kind patterns, plus
    // each three-item sequence with a plain paragraph after the first item.
    static IEnumerable<ListLine?[]> ListLineSequences()
    {
        for (var length = 1; length <= 4; length++)
        {
            var count = (int) Math.Pow(4, length);
            for (var code = 0; code < count; code++)
            {
                var levels = new int[length];
                var rest = code;
                for (var i = 0; i < length; i++)
                {
                    levels[i] = rest % 4;
                    rest /= 4;
                }

                yield return levels.Select(_ => new ListLine(_)).ToArray<ListLine?>();
                yield return levels.Select((_, i) => new ListLine(_, i % 2 == 1)).ToArray<ListLine?>();
                yield return levels.Select((_, i) => new ListLine(_, i >= 2)).ToArray<ListLine?>();

                if (length == 3)
                {
                    yield return [new(levels[0]), null, new(levels[1], true), new(levels[2])];
                }
            }
        }
    }

    static string DescribeLine(ListLine? line, int index)
    {
        if (line == null)
        {
            return $"p:i{index}";
        }

        if (line.Ordered)
        {
            return $"ol{line.Level}:i{index}";
        }

        return $"ul{line.Level}:i{index}";
    }

    // The same notation as Describe, read back from the html: each element that directly holds
    // text, with the kind of its list and how many lists deep it sits.
    static string DescribeHtml(XElement root)
    {
        var parts = new List<string>();
        foreach (var element in root.Descendants())
        {
            var text = string.Concat(element.Nodes().OfType<XText>().Select(_ => _.Value));
            if (text.Length == 0)
            {
                continue;
            }

            if (element.Name.LocalName == "p")
            {
                parts.Add($"p:{text}");
                continue;
            }

            var depth = element.Ancestors().Count(_ => _.Name.LocalName is "ul" or "ol") - 1;
            parts.Add($"{element.Parent!.Name.LocalName}{depth}:{text}");
        }

        return string.Join(",", parts);
    }

    static async Task<string?> ExtractListLines(params ListLine?[] lines)
    {
        using var rendered = await RenderModel(
            "{{ Body }}",
            new EditableArticle
            {
                Title = "T",
                Body = listSeed
            });
        return ExtractListLinesFrom(rendered, lines);
    }

    // Replaces the rendered control's content with the given paragraphs, the way a user's Tab and
    // Shift+Tab leave them, and extracts the field.
    static string? ExtractListLinesFrom(MemoryStream rendered, ListLine?[] lines)
    {
        using var copy = new MemoryStream();
        rendered.Position = 0;
        rendered.CopyTo(copy);

        using (var doc = WordprocessingDocument.Open(copy, true))
        {
            var sdt = FindSdtBlock(doc.MainDocumentPart!.Document!.Body!, "Body");
            var content = sdt.ChildElements.First(_ => _.LocalName == "sdtContent");
            var numberIds = content
                .Descendants<NumberingId>()
                .Select(_ => _.Val!.Value)
                .ToList();
            content.RemoveAllChildren();
            for (var i = 0; i < lines.Length; i++)
            {
                var paragraph = new Paragraph();
                var line = lines[i];
                if (line != null)
                {
                    var numberId = numberIds[0];
                    if (line.Ordered)
                    {
                        numberId = numberIds[1];
                    }

                    paragraph.ParagraphProperties = new()
                    {
                        NumberingProperties = new()
                        {
                            NumberingLevelReference = new()
                            {
                                Val = line.Level
                            },
                            NumberingId = new()
                            {
                                Val = numberId
                            }
                        }
                    };
                }

                paragraph.AppendChild(new Run(new Text($"i{i}")));
                content.AppendChild(paragraph);
            }

            doc.Save();
        }

        copy.Position = 0;
        var model = new EditableArticle
        {
            Title = "T",
            Body = ""
        };
        ParchmentExtractor.Extract<EditableArticle>(copy).ApplyTo(model);
        return model.Body;
    }
}
