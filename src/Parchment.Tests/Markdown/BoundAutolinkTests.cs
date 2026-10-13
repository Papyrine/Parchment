// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// A url written in the template with a bound tail. An autolink is read verbatim, so the escapes
/// MarkdownEncoder writes into the value have to be undone for the link to work — without the value
/// gaining the means to end the link and start markup of its own.
/// </summary>
public partial class BoundAutolinkTests
{
    const string Prefix = "https://portal.example.com/orders/";

    [ParchmentBindable]
    public partial class LinkModel
    {
        public string Id { get; init; } = "";
    }

    [Test]
    public async Task BareUrlWithBoundTail()
    {
        var (body, targets) = await Render(
            "See https://portal.example.com/orders/{{ Id }} for details.",
            "A-1_x.y");

        await Assert.That(targets).IsEquivalentTo([Prefix + "A-1_x.y"]);
        await Assert.That(body.Descendants<Hyperlink>().Single().InnerText).IsEqualTo(Prefix + "A-1_x.y");
        await Assert.That(body.InnerText).IsEqualTo($"See {Prefix}A-1_x.y for details.");
    }

    [Test]
    public async Task AngleAutolinkWithBoundTail()
    {
        var (body, targets) = await Render(
            "<https://portal.example.com/orders/{{ Id }}>",
            "A-1_x.y");

        await Assert.That(targets).IsEquivalentTo([Prefix + "A-1_x.y"]);
        await Assert.That(body.InnerText).IsEqualTo(Prefix + "A-1_x.y");
    }

    [Test]
    public async Task WwwUrlWithBoundTail()
    {
        var (body, targets) = await Render(
            "See www.example.com/orders/{{ Id }} for details.",
            "A-1_x.y");

        await Assert.That(targets).IsEquivalentTo(["http://www.example.com/orders/A-1_x.y"]);
        await Assert.That(body.Descendants<Hyperlink>().Single().InnerText).IsEqualTo("www.example.com/orders/A-1_x.y");
    }

    // An inline link's destination goes through CommonMark's own escape processing, so it never
    // needed the help.
    [Test]
    [Arguments("A-1_x.y")]
    [Arguments("a(b)")]
    [Arguments("x](http://evil)")]
    [Arguments("q?x=1&y=2")]
    [Arguments("end)")]
    public async Task InlineLinkDestinationWithBoundTail(string value)
    {
        var (body, targets) = await Render(
            "[text](https://portal.example.com/orders/{{ Id }})",
            value);

        await Assert.That(targets).IsEquivalentTo([Prefix + value]);
        await Assert.That(body.InnerText).IsEqualTo("text");
    }

    // A space or a control character is the one thing an escape cannot keep in a destination: it
    // ended the destination, the link was no longer a link, and the markdown printed as text. The
    // target is still the value as bound, not a percent-encoded copy of it.
    [Test]
    [Arguments("a b")]
    [Arguments(" a  b ")]
    [Arguments("a\tb")]
    [Arguments("a b (1)")]
    [Arguments("a) b")]
    [Arguments("a <b> c")]
    [Arguments("a \"b\" 'c'")]
    [Arguments("a b](http://evil)")]
    [Arguments("a b) *c*")]
    public async Task InlineLinkDestinationWithASpaceInTheValue(string value)
    {
        var (body, targets) = await Render(
            "[text](https://portal.example.com/orders/{{ Id }})",
            value);

        await Assert.That(targets).IsEquivalentTo([Prefix + value]);
        await Assert.That(body.InnerText).IsEqualTo("text");
        await AssertNoMarkup(body);
    }

    [Test]
    [Arguments("[text]({{ Id }})")]
    [Arguments("[text](<{{ Id }}>)")]
    [Arguments("[text]({{ Id }} \"a title\")")]
    [Arguments("[text]({{ Id }} 'a title')")]
    [Arguments("[text]({{ Id }} (a title))")]
    [Arguments("[text](<{{ Id }}> \"a title\")")]
    [Arguments("{% if true %}[text]({% endif %}{{ Id }})")]
    [Arguments("| [text]({{ Id }}) |\n| --- |")]
    [Arguments("- [text]({{ Id }})")]
    [Arguments("> **[text]({{ Id }})**")]
    public async Task WholeDestinationBoundToAValueWithASpace(string template)
    {
        var value = "https://example.com/a b (1)/c d?q=\"x y\"";
        var (body, targets) = await Render(template, value);

        await Assert.That(targets).IsEquivalentTo([value]);
        await Assert.That(body.InnerText).IsEqualTo("text");
    }

    [Test]
    public async Task TwoValuesInOneDestination()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<LinkModel>(
            "[one](https://portal.example.com/{{ Id }}/{{ Id }}) and [two]({{ Id }}) for {{ Id }}.",
            styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new LinkModel
            {
                Id = "a b"
            },
            stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var targets = main.HyperlinkRelationships
            .Select(_ => _.Uri.OriginalString)
            .ToList();

        await Assert.That(targets).IsEquivalentTo(["https://portal.example.com/a b/a b", "a b"]);
        await Assert.That(main.Document!.Body!.InnerText).IsEqualTo("one and two for a b.");
    }

    // The same value outside a destination is unchanged: a closed destination is not one, and
    // neither is a bracket the template escaped. A code span cannot be told apart from the output
    // alone, so link syntax quoted in one shows the character references.
    [Test]
    [Arguments("[text](https://example.com/x){{ Id }}", "texta  b")]
    [Arguments("[text](https://example.com/x(y)){{ Id }}", "texta  b")]
    [Arguments(@"\[text\]({{ Id }})", "[text](a  b)")]
    [Arguments("`[text]({{ Id }})`", "[text](a&#32;&#32;b)")]
    public async Task ValueWithASpaceOutsideADestination(string template, string text)
    {
        var (body, _) = await Render(template, "a  b");

        await Assert.That(body.InnerText).IsEqualTo(text);
    }

    // There is no line break in a url. The value still reaches the link whole rather than ending
    // it with the html a break is written as everywhere else.
    [Test]
    [Arguments("a\nb")]
    [Arguments("a\r\nb")]
    public async Task NewlineInADestinationValue(string value)
    {
        var (body, targets) = await Render(
            "[text](https://portal.example.com/orders/{{ Id }})",
            value);

        await Assert.That(targets).IsEquivalentTo([Prefix + value]);
        await Assert.That(body.InnerText).IsEqualTo("text");
        await Assert.That(body.Descendants<Break>().Any()).IsFalse();
    }

    // No destination is no link, and the text is still there.
    [Test]
    [Arguments("[text]({{ Id }})", "text")]
    [Arguments("![alt]({{ Id }})", "alt")]
    public async Task EmptyDestinationValue(string template, string text)
    {
        var (body, targets) = await Render(template, "");

        await Assert.That(targets).IsEmpty();
        await Assert.That(body.InnerText).IsEqualTo(text);
    }

    // An image source is a path read from disk, so it has to arrive with its spaces and brackets
    // as they are on disk. It used to print — the path, in the document — instead of loading.
    [Test]
    [Arguments("![pixel]({{ Id }})", false)]
    [Arguments("![pixel](<{{ Id }}>)", false)]
    [Arguments("![pixel]({{ Id }} \"a title\")", false)]
    [Arguments("![pixel]({{ Id }})", true)]
    public async Task ImageSourceWithASpaceInTheValue(string template, bool forwardSlashes)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"parchment md img (1) {Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "pixel two.png");
        await File.WriteAllBytesAsync(
            path,
            Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORk5CYII="));
        try
        {
            if (forwardSlashes)
            {
                path = path.Replace('\\', '/');
            }

            var (body, _) = await Render(template, path);

            await Assert.That(body.Descendants<Drawing>().Count()).IsEqualTo(1);
            await Assert.That(body.InnerText).IsEqualTo("");
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    // Every character of the value prints exactly once and none of it is markup. Where the link
    // ends is Markdig's call, the same one it makes for a url typed out in full: not at a space, a
    // < or an entity, and not on trailing punctuation.
    [Test]
    [Arguments("q?x=1&y=2", "q?x=1&y=2")]
    [Arguments("a(b)", "a(b)")]
    [Arguments(@"a\b", @"a\b")]
    [Arguments(@"a\", @"a\")]
    [Arguments("a>b", "a>b")]
    [Arguments("x](http://evil)", "x](http://evil)")]
    [Arguments("a b", "a")]
    [Arguments("end.", "end")]
    [Arguments("end)", "end")]
    [Arguments("*x*", "*x")]
    [Arguments("a*", "a")]
    [Arguments("a_", "a")]
    [Arguments("**x** y", "**x*")]
    [Arguments("<script>", "")]
    [Arguments("<b>bold</b>", "")]
    [Arguments("a&amp;b", "a")]
    [Arguments("a<http://evil.example>", "a")]
    [Arguments("a [evil](http://evil.example)", "a")]
    public async Task HostileValueInABareUrl(string value, string linkedTail)
    {
        var (body, targets) = await Render(
            "See https://portal.example.com/orders/{{ Id }} for details.",
            value);

        await Assert.That(body.InnerText).IsEqualTo($"See {Prefix}{value} for details.");
        await Assert.That(targets).IsEquivalentTo([Prefix + linkedTail]);
        await Assert.That(body.Descendants<Hyperlink>().Single().InnerText).IsEqualTo(Prefix + linkedTail);
        await AssertNoMarkup(body);
    }

    // The character after the link is where the escape matters most: the link took the backslash
    // with it, which would leave a * free to close the template's own emphasis early.
    [Test]
    public async Task ValueCannotCloseTheTemplatesEmphasis()
    {
        var (body, targets) = await Render(
            "*see https://portal.example.com/orders/{{ Id }} now*",
            "a*");

        await Assert.That(body.InnerText).IsEqualTo($"see {Prefix}a* now");
        await Assert.That(targets).IsEquivalentTo([Prefix + "a"]);
        var tail = body.Descendants<Run>().Last();
        await Assert.That(tail.InnerText).EndsWith(" now");
        await Assert.That(tail.RunProperties?.Italic).IsNotNull();
    }

    [Test]
    [Arguments("x](http://evil)", "x](http://evil)", "")]
    [Arguments("*x*", "*x*", "")]
    [Arguments("end.", "end.", "")]
    [Arguments("end)", "end)", "")]
    [Arguments(@"a\b", @"a\b", "")]
    [Arguments(@"a\", @"a\", "")]
    // The value's own > closes the autolink, so the rest of it — and the template's > — is text.
    [Arguments("a>b", "a", ">b>")]
    [Arguments("a>*b*", "a", ">*b*>")]
    public async Task HostileValueInAnAngleAutolink(string value, string linkedTail, string after)
    {
        var (body, targets) = await Render(
            "<https://portal.example.com/orders/{{ Id }}>",
            value);

        await Assert.That(body.InnerText).IsEqualTo(Prefix + linkedTail + after);
        await Assert.That(targets).IsEquivalentTo([Prefix + linkedTail]);
        await Assert.That(body.Descendants<Hyperlink>().Single().InnerText).IsEqualTo(Prefix + linkedTail);
        await AssertNoMarkup(body);
    }

    // A space or a < cannot be in an autolink at all, so there is no link and the whole line is
    // text.
    [Test]
    [Arguments("a b")]
    [Arguments("<script>")]
    public async Task ValueThatVoidsAnAngleAutolink(string value)
    {
        var (body, targets) = await Render(
            "<https://portal.example.com/orders/{{ Id }}>",
            value);

        await Assert.That(body.InnerText).IsEqualTo($"<{Prefix}{value}>");
        await Assert.That(targets).IsEmpty();
        await AssertNoMarkup(body);
    }

    [Test]
    [Arguments("See https://portal.example.com/orders/{{ Id }} for details.", "See " + Prefix + "ab for details.")]
    [Arguments("<https://portal.example.com/orders/{{ Id }}>", "<" + Prefix + "ab>")]
    public async Task NewlineInAValueEndsTheLinkWithABreak(string template, string text)
    {
        var (body, _) = await Render(template, "a\nb");

        await Assert.That(body.Elements<Paragraph>().Count()).IsEqualTo(1);
        await Assert.That(body.Descendants<Break>().Count()).IsEqualTo(1);
        await Assert.That(body.InnerText).IsEqualTo(text);
        await Assert.That(body.Descendants<Hyperlink>().Any(_ => _.InnerText.Contains('b'))).IsFalse();
    }

    // A url that arrives whole inside a value is text: nothing in the template said it was a link.
    [Test]
    public async Task UrlInsideAValueIsNotLinked()
    {
        var value = "see https://example.com/my-page.html or www.example.com/a_b or mail bob.smith@example.com ok";
        var (body, targets) = await Render("{{ Id }}", value);

        await Assert.That(body.InnerText).IsEqualTo(value);
        await Assert.That(targets).IsEmpty();
    }

    static async Task AssertNoMarkup(Body body)
    {
        await Assert.That(body.Descendants<Italic>().Any()).IsFalse();
        await Assert.That(body.Descendants<Bold>().Any()).IsFalse();
        await Assert.That(body.Elements<Paragraph>().Count()).IsEqualTo(1);
    }

    static async Task<(Body body, List<string> targets)> Render(string template, string id)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<LinkModel>(template, styleSource);

        using var stream = new MemoryStream();
        await store.Render(
            new LinkModel
            {
                Id = id
            },
            stream);
        stream.Position = 0;

        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var targets = main.HyperlinkRelationships
            .Select(_ => _.Uri.OriginalString)
            .ToList();
        // Cloned so it outlives the package.
        return ((Body) main.Document!.Body!.CloneNode(true), targets);
    }
}
