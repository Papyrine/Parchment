using DocumentFormat.OpenXml.Validation;

/// <summary>
/// An entity is the character it names. Nothing rendered one before, and a leaf inline with no
/// renderer writes nothing at all, so <c>AT&amp;amp;T</c> reached the document as "ATT".
/// </summary>
public class HtmlEntityInlineRendererTests
{
    [Test]
    public async Task EmitsTheDecodedCharacter()
    {
        var entity = RendererHarness.FirstInline<HtmlEntityInline>("AT&amp;T");
        var renderer = RendererHarness.BuildRenderer();

        renderer.Render(entity);

        var run = (Run) renderer.Top.CurrentRuns.Single();
        var text = run.GetFirstChild<Text>()!;
        await Assert.That(text.Text).IsEqualTo("&");
        await Assert.That(text.Space?.Value).IsEqualTo(SpaceProcessingModeValues.Preserve);
    }

    [Test]
    [Arguments("AT&amp;T and R&amp;D", "AT&T and R&D")]
    [Arguments("&copy; 2026 Acme", "\u00A9 2026 Acme")]
    [Arguments("use &lt;tag&gt; here", "use <tag> here")]
    [Arguments("&#169; and &#xA9; and &#XA9;", "\u00A9 and \u00A9 and \u00A9")]
    [Arguments("10&nbsp;km", "10\u00A0km")]
    [Arguments("&quot;x&quot; and &apos;y&apos;", "\"x\" and 'y'")]
    [Arguments("a &mdash; b &hellip;", "a \u2014 b \u2026")]
    // One entity, two characters.
    [Arguments("a &NotEqualTilde; b", "a \u2242\u0338 b")]
    // Outside the BMP, so a surrogate pair.
    [Arguments("&#x1F600;", "\U0001F600")]
    [Arguments("&amp;", "&")]
    // The entity is decoded once: what it decodes to is text, not another entity.
    [Arguments("&amp;amp; and &amp;lt;", "&amp; and &lt;")]
    public async Task EntityRendersAsTheCharacterItNames(string markdown, string expected)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo(expected);

        // Plain text either side of an entity is the same formatting, so it is still one run.
        await Assert.That(paragraph.Elements<Run>().Count()).IsEqualTo(1);
        await Assert.That(Validate(paragraph)).IsEmpty();
    }

    // What is not an entity was never one, and stays as typed.
    [Test]
    [Arguments("&notanentity; stays", "&notanentity; stays")]
    [Arguments("fish &amp chips", "fish &amp chips")]
    [Arguments("R&D", "R&D")]
    [Arguments("`a &amp; b`", "a &amp; b")]
    [Arguments(@"\&amp;", "&amp;")]
    public async Task TextThatIsNotAnEntityStaysAsTyped(string markdown, string expected)
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown(markdown).Single();

        await Assert.That(paragraph.InnerText).IsEqualTo(expected);
    }

    // A tab is a tab however it was spelled, so it takes the same route a literal one does.
    [Test]
    public async Task TabEntityBecomesAWordTab()
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown("A&#9;B").Single();

        await Assert.That(paragraph.Descendants<TabChar>().Count()).IsEqualTo(1);
        await Assert.That(paragraph.InnerText).IsEqualTo("AB");
        await Assert.That(Validate(paragraph)).IsEmpty();
    }

    // A character xml cannot carry is dropped the way it is from literal text, rather than written
    // into a package Word then refuses to open.
    [Test]
    public async Task CharacterXmlCannotCarryIsDropped()
    {
        var paragraph = (Paragraph) RendererHarness.RenderMarkdown("a&#1;b").Single();

        await Assert.That(paragraph.InnerText).IsEqualTo("ab");
        await Assert.That(paragraph.Elements<Run>().Count()).IsEqualTo(1);
    }

    static IEnumerable<string> Validate(OpenXmlElement element) =>
        new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(element)
            .Select(_ => $"{_.Description} @ {_.Path?.XPath}");
}
