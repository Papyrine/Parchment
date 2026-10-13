/// <summary>
/// A literal is folded into the run of the literal before it while the two are adjacent siblings,
/// instead of getting a run that the paragraph flush folds away again. These pin the run boundaries
/// that have to come out the same either way: the ones the flush folds, and the ones it never did.
/// </summary>
public class LiteralFoldingTests
{
    // Each paragraph is its children separated by spaces: a run is its rPr children in brackets,
    // then its content, and a hyperlink is link(...) around its runs.
    [Test]
    // Every escape starts another literal; all of them are one run.
    [Arguments(@"on\-call re\-run \(twice\)\.", "'on-call re-run (twice).'")]
    // Formatting is still applied to the whole of what an emphasis holds, and to nothing else.
    [Arguments(@"*a\-b* c\-d", "[i]'a-b' ' c-d'")]
    [Arguments(@"a\-b *c\-d*", "'a-b ' [i]'c-d'")]
    [Arguments(@"a\-b*c*\-d\-e", "'a-b' [i]'c' '-d-e'")]
    [Arguments(@"***a\-b** c\-d* e\-f", "[b,i]'a-b' [i]' c-d' ' e-f'")]
    [Arguments(@"**a\-b**{.Lead} c\-d", "[rStyle,b]'a-b' ' c-d'")]
    // The runs of a link's text go into the hyperlink, which is not folded.
    [Arguments(@"[a\-b](https://example.com) c\-d", "link([rStyle]'a' [rStyle]'-b') ' c-d'")]
    [Arguments(@"x\-y [a\-b](https://example.com)", "'x-y ' link([rStyle]'a' [rStyle]'-b')")]
    [Arguments(@"**[a\-b](https://example.com)** c\-d", "link([rStyle,b]'a' [rStyle,b]'-b') ' c-d'")]
    [Arguments(@"[*a\-b* c\-d](https://example.com)", "link([rStyle,i]'a' [rStyle,i]'-b' [rStyle]' c' [rStyle]'-d')")]
    [Arguments(@"[a\-b](#place) c\-d", "link([rStyle]'a' [rStyle]'-b') ' c-d'")]
    // A link with nowhere to point stays in the paragraph as text, and is folded with it.
    [Arguments(@"x\-y [a\-b](../relative) c\-d", "'x-y a-b (../relative) c-d'")]
    // Inline html formats what is written while the tag is open.
    [Arguments(@"<b>a\-b</b>c\-d", "[b]'a-b' 'c-d'")]
    [Arguments(@"a\-b<u>c\-d", "'a-b' [u]'c-d'")]
    [Arguments(@"a\-b<span class=""Mark"">c\-d</span>\-e", "'a-b' [rStyle]'c-d' '-e'")]
    [Arguments(@"<i>a\-b *c\-d*</i>", "[i]'a-b c-d'")]
    // A tab makes a run of several pieces, which is folded with nothing.
    [Arguments("a\\-b\tc\\-d\\-e", "'a' '-b'<tab>'c' '-d-e'")]
    [Arguments("\\-a\\-b\tc", "'-a' '-b'<tab>'c'")]
    // Text from the other inlines is still folded at the flush.
    [Arguments(@"a\-b&amp;c\-d", "'a-b&c-d'")]
    [Arguments("a\\-b\nc\\-d", "'a-b c-d'")]
    [Arguments(@"a\-b ""c\-d""", "'a-b “c-d”'")]
    [Arguments(@"see [a\-b c\-d", "'see [a-b c-d'")]
    // And what is not text keeps its place between the runs.
    [Arguments("a\\-b\\\nc\\-d", "'a-b' <br> 'c-d'")]
    [Arguments(@"a\-b<br>c\-d", "'a-b' <br> 'c-d'")]
    [Arguments(@"a\-b`c\-d`e\-f", "'a-b' [rFonts]'c\\-d' 'e-f'")]
    [Arguments(@"a\-b <https://example.com> c\-d", "'a-b ' link([rStyle]'https://example.com') ' c-d'")]
    // Whitespace at either end of a piece is kept.
    [Arguments(@"a \-  b\- *c*", "'a -  b- ' [i]'c'")]
    // A character xml cannot carry is dropped from whichever literal held it.
    [Arguments("a\\-\u0001\\-b\\-\u0001", "'a--b-'")]
    // Each block starts again.
    [Arguments("a\\-b\n\nc\\-d", "'a-b' / 'c-d'")]
    [Arguments("- a\\-b\n- c\\-d", "'a-b' / 'c-d'")]
    [Arguments("# a\\-b {#place}\n\nc\\-d", "<bookmarkStart> 'a-b' <bookmarkEnd> / 'c-d'")]
    [Arguments("| a\\-b | *c\\-d* e\\-f |\n| --- | --- |\n| g\\-h | [i\\-j](https://example.com) |", "[b]'a-b' / [b,i]'c-d' [b]' e-f' / 'g-h' / link([rStyle]'i' [rStyle]'-j')")]
    public async Task RunBoundaries(string markdown, string expected)
    {
        var elements = RendererHarness.RenderMarkdown(markdown);

        var paragraphs = elements
            .SelectMany(_ => _.Descendants<Paragraph>().Prepend(_).OfType<Paragraph>())
            .Select(Describe);
        await Assert.That(string.Join(" / ", paragraphs)).IsEqualTo(expected);

        foreach (var text in elements.SelectMany(_ => _.Descendants<Text>()))
        {
            await Assert.That(text.Space?.Value).IsEqualTo(SpaceProcessingModeValues.Preserve);
        }
    }

    // A table cell is a container of its own, and what is written into it must not reach back
    // into the run that was last in the container around it.
    [Test]
    public async Task LiteralIsNotFoldedIntoARunOfAnotherContainer()
    {
        var renderer = RendererHarness.BuildRenderer();
        var paragraph = RendererHarness.FirstBlock<ParagraphBlock>(@"a\-b");
        var first = paragraph.Inline!.FirstChild!;
        var second = first.NextSibling!;

        renderer.Render(first);
        renderer.PushContainer();
        renderer.Render(second);

        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(1);
        await Assert.That(renderer.Top.CurrentRuns[0].InnerText).IsEqualTo("-b");
        renderer.PopContainer();
        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(1);
        await Assert.That(renderer.Top.CurrentRuns[0].InnerText).IsEqualTo("a");
    }

    [Test]
    public async Task LiteralIsNotFoldedOnceItsRunIsNoLongerTheLast()
    {
        var renderer = RendererHarness.BuildRenderer();
        var paragraph = RendererHarness.FirstBlock<ParagraphBlock>(@"a\-b");
        var first = paragraph.Inline!.FirstChild!;

        renderer.Render(first);
        renderer.AddRun(new Run(new Break()));
        renderer.Render(first.NextSibling!);

        await Assert.That(renderer.Top.CurrentRuns.Count).IsEqualTo(3);
        await Assert.That(renderer.Top.CurrentRuns[0].InnerText).IsEqualTo("a");
        await Assert.That(renderer.Top.CurrentRuns[2].InnerText).IsEqualTo("-b");
    }

    static string Describe(Paragraph paragraph) =>
        string.Join(' ', paragraph.ChildElements.Where(_ => _ is not ParagraphProperties).Select(Describe));

    static string Describe(OpenXmlElement element)
    {
        if (element is Hyperlink hyperlink)
        {
            return $"link({string.Join(' ', hyperlink.ChildElements.Select(Describe))})";
        }

        if (element is not Run run)
        {
            return $"<{element.LocalName}>";
        }

        var builder = new StringBuilder();
        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case RunProperties properties:
                    builder.Append($"[{string.Join(',', properties.ChildElements.Select(_ => _.LocalName))}]");
                    break;
                case Text text:
                    builder.Append($"'{text.Text}'");
                    break;
                default:
                    builder.Append($"<{child.LocalName}>");
                    break;
            }
        }

        return builder.ToString();
    }
}
