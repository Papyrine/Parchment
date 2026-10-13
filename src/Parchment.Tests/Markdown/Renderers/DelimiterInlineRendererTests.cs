using Markdig.Parsers.Inlines;

/// <summary>
/// A delimiter that never found its match stays in the tree as a delimiter, holding everything
/// after it. It is the text the author typed: <c>see [note</c> has a bracket in it.
/// </summary>
public class DelimiterInlineRendererTests
{
    [Test]
    [Arguments("see [note for details", "see [note for details")]
    [Arguments("see ![note for details", "see ![note for details")]
    [Arguments("[outer [inner", "[outer [inner")]
    [Arguments("[", "[")]
    [Arguments("[ref][nope] and [text](", "[ref][nope] and [text](")]
    [Arguments("a [b **c** d", "a [b c d")]
    [Arguments("a [b &amp; <<c>>", "a [b & \u00ABc\u00BB")]
    // A bracket that does close, into a link or into nothing, was never left as a delimiter.
    [Arguments("a ] b", "a ] b")]
    [Arguments("[not a link] here", "[not a link] here")]
    public async Task UnmatchedDelimiterRendersAsTyped(string markdown, string expected)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo(expected);
    }

    // A pipe outside a table is a delimiter too, and is only turned back into text when it sits
    // directly in the paragraph. Inside emphasis, or after another delimiter, it stays one.
    [Test]
    [Arguments("a | b", "a | b")]
    [Arguments("*a | b*", "a | b")]
    [Arguments("**yes | no** and ~~x | y~~", "yes | no and x | y")]
    [Arguments("a [b | c", "a [b | c")]
    [Arguments("a | b [c", "a | b [c")]
    [Arguments("| lone [ |", "| lone [ |")]
    [Arguments("`a | b` | [c", "a | b | [c")]
    [Arguments("a | b\n[c | d", "a | b [c | d")]
    public async Task PipeOutsideATableRendersAsTyped(string markdown, string expected)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo(expected);
    }

    // The bracket does not stop the row it is in from being a row.
    [Test]
    public async Task UnmatchedBracketInATableCellStaysInItsCell()
    {
        var table = (Table) RendererHarness.RenderMarkdown("| a [b | c |\n| --- | --- |\n| d | *e [f* |").Single();

        await Assert.That(table.Descendants<TableCell>().Select(_ => _.InnerText)).IsEquivalentTo(["a [b", "c", "d", "e [f"]);
    }

    // The open bracket takes the rest of the block as its children, formatting included.
    [Test]
    public async Task ContentAfterTheDelimiterKeepsItsFormatting()
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown("a [b **c** d").Single();

        var runs = paragraph.Elements<Run>()
            .Select(_ => $"{string.Join(' ', _.RunProperties?.Select(property => property.LocalName) ?? [])}|{_.InnerText}")
            .ToList();

        await Assert.That(runs).IsEquivalentTo(["|a [b ", "b|c", "| d"]);
    }

    // A link after the stray bracket is still a link, and the bracket still prints.
    [Test]
    public async Task LinkAfterAnUnmatchedBracketIsStillALink()
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown("[a [b](https://example.com) c").Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("[a b c");
        await Assert.That(paragraph.Elements<Hyperlink>().Single().InnerText).IsEqualTo("b");
    }

    // The renderer is registered for the base type, so a delimiter kind left behind by any parser
    // prints its own text and then what it holds - not only the link parser's bracket.
    [Test]
    public async Task AnyDelimiterKindWritesItsLiteralThenItsChildren()
    {
        var delimiter = new StubDelimiter("||");
        delimiter.AppendChild(new LiteralInline("after"));
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(delimiter);

        await Assert.That(renderer.Top.CurrentRuns.Select(_ => _.InnerText)).IsEquivalentTo(["||", "after"]);
    }

    [Test]
    public async Task EmptyLiteralWritesOnlyTheChildren()
    {
        var delimiter = new StubDelimiter(string.Empty);
        delimiter.AppendChild(new LiteralInline("after"));
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(delimiter);

        await Assert.That(renderer.Top.CurrentRuns.Select(_ => _.InnerText)).IsEquivalentTo(["after"]);
    }

    class StubDelimiter(string literal) :
        DelimiterInline(new LiteralInlineParser())
    {
        public override string ToLiteral() =>
            literal;
    }
}
