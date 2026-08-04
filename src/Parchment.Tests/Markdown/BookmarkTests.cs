public class BookmarkTests
{
    [Test]
    public async Task HeadingIdBecomesBookmarkAroundTheText()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("## Section {#intro}").Single();

        var start = paragraph.GetFirstChild<BookmarkStart>()!;
        await Assert.That(start.Name?.Value).IsEqualTo("intro");

        // Wrapping the runs rather than sitting beside them is what gives a PAGEREF something to
        // resolve a page from.
        var children = paragraph.ChildElements.ToList();
        await Assert.That(children.IndexOf(start)).IsLessThan(children.FindIndex(_ => _ is Run));
        await Assert.That(paragraph.GetFirstChild<BookmarkEnd>()!.Id?.Value).IsEqualTo(start.Id?.Value);
    }

    [Test]
    public async Task ParagraphIdBecomesBookmark()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("Body text. {#note}").Single();

        await Assert.That(paragraph.GetFirstChild<BookmarkStart>()!.Name?.Value).IsEqualTo("note");
    }

    [Test]
    public async Task StyleAndIdBothApply()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("Body text. {#note}{.BOXText}").Single();

        await Assert.That(paragraph.ParagraphProperties!.ParagraphStyleId!.Val?.Value).IsEqualTo("BOXText");
        await Assert.That(paragraph.GetFirstChild<BookmarkStart>()!.Name?.Value).IsEqualTo("note");
    }

    [Test]
    public async Task NoIdEmitsNoBookmark()
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown("Plain paragraph.").Single();

        await Assert.That(paragraph.GetFirstChild<BookmarkStart>()).IsNull();
    }

    // A bookmark name is far more restricted than an html id, and rejecting one markdown accepts
    // would make a template that also renders to html unusable — so it is folded into shape. A
    // space is not among the cases: it ends the attribute in markdown's own syntax, so an id
    // containing one never reaches here.
    [Test]
    [Arguments("a-b", "a_b")]
    [Arguments("a.b:c", "a_b_c")]
    [Arguments("_leading", "_leading")]
    [Arguments("9lives", "_9lives")]
    [Arguments("-dash", "_dash")]
    public async Task NameIsFoldedToWordRules(string id, string expected)
    {
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown($"Text {{#{id}}}").Single();

        await Assert.That(paragraph.GetFirstChild<BookmarkStart>()!.Name?.Value).IsEqualTo(expected);
    }

    [Test]
    public async Task LongNameIsTruncatedToFortyCharacters()
    {
        var id = new string('a', 60);
        var paragraph = (Paragraph)RendererHarness.RenderMarkdown($"Text {{#{id}}}").Single();

        await Assert.That(paragraph.GetFirstChild<BookmarkStart>()!.Name?.Value).IsEqualTo(new string('a', 40));
    }

    // Word pairs a bookmark's start to its end by id, so two bookmarks sharing one id close the
    // wrong range. The html converter numbers its own bookmarks from 1 for every block it converts,
    // which is exactly the collision the renumbering pass exists to prevent.
    [Test]
    public async Task IdsAreUniqueAcrossMarkdownAndHtmlBookmarks()
    {
        var elements = RendererHarness.RenderMarkdown(
            """
            # First {#one}

            <p id="two">From html</p>

            <p id="three">Also html</p>

            # Last {#four}
            """);

        var starts = elements.SelectMany(_ => _.Descendants<BookmarkStart>().Concat(AsBookmark(_))).ToList();
        var ids = starts.Select(_ => _.Id?.Value).ToList();

        await Assert.That(starts.Count).IsEqualTo(4);
        await Assert.That(ids.Distinct().Count()).IsEqualTo(4);
    }

    static IEnumerable<BookmarkStart> AsBookmark(OpenXmlElement element)
    {
        if (element is BookmarkStart start)
        {
            yield return start;
        }
    }
}
