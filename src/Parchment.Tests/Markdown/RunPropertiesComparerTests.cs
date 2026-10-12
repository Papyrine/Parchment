/// <summary>
/// The comparer stands in for comparing the <c>OuterXml</c> of two <c>w:rPr</c>, which is what
/// decided whether neighbouring runs fold into one. Every case here is asserted against that
/// comparison as well as against the expected answer, so the two cannot drift apart.
/// </summary>
public class RunPropertiesComparerTests
{
    const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string word14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    [Test]
    public Task BothMissing() =>
        Check(null, null, true);

    [Test]
    public async Task MissingIsNotEmpty()
    {
        await Check(null, new(), false);
        await Check(new(), null, false);
    }

    [Test]
    public Task BothEmpty() =>
        Check(new(), new(), true);

    [Test]
    public Task SameInstance()
    {
        var properties = new RunProperties(new Bold());
        return Check(properties, properties, true);
    }

    [Test]
    public Task EmptyIsNotBold() =>
        Check(new(), new(new Bold()), false);

    [Test]
    public Task SameElement() =>
        Check(new(new Bold()), new(new Bold()), true);

    [Test]
    public Task DifferentElement() =>
        Check(new(new Bold()), new(new Italic()), false);

    [Test]
    public async Task DifferentChildCount()
    {
        await Check(new(new Bold()), new(new Bold(), new Italic()), false);
        await Check(new(new Bold(), new Italic()), new(new Bold()), false);
    }

    [Test]
    public Task DifferentChildOrder() =>
        Check(new(new Bold(), new Italic()), new(new Italic(), new Bold()), false);

    [Test]
    public async Task BoldWithAndWithoutValue()
    {
        await Check(new(new Bold()), new(new Bold { Val = true }), false);
        await Check(new(new Bold { Val = true }), new(new Bold { Val = true }), true);
        await Check(new(new Bold { Val = true }), new(new Bold { Val = false }), false);
    }

    [Test]
    public async Task RunStyle()
    {
        await Check(Styled("Hyperlink"), Styled("Hyperlink"), true);
        await Check(Styled("Hyperlink"), Styled("Lead"), false);
        await Check(Styled("Hyperlink"), new(new RunStyle()), false);
    }

    [Test]
    public async Task RunFontsWithSeveralAttributes()
    {
        await Check(Fonts("Consolas", "Consolas"), Fonts("Consolas", "Consolas"), true);
        await Check(Fonts("Consolas", "Consolas"), Fonts("Consolas", "Courier New"), false);
        await Check(
            Fonts("Consolas", "Consolas"),
            new(
                new RunFonts
                {
                    Ascii = "Consolas"
                }),
            false);
        await Check(
            new(
                new RunFonts
                {
                    Ascii = "Consolas"
                }),
            new(
                new RunFonts
                {
                    HighAnsi = "Consolas"
                }),
            false);
    }

    [Test]
    public async Task Color()
    {
        await Check(Colored("FF0000"), Colored("FF0000"), true);
        await Check(Colored("FF0000"), Colored("00FF00"), false);
    }

    [Test]
    public async Task Highlight()
    {
        await Check(Highlighted(HighlightColorValues.Yellow), Highlighted(HighlightColorValues.Yellow), true);
        await Check(Highlighted(HighlightColorValues.Yellow), Highlighted(HighlightColorValues.Green), false);
    }

    [Test]
    public async Task Underline()
    {
        await Check(Underlined(UnderlineValues.Single), Underlined(UnderlineValues.Single), true);
        await Check(Underlined(UnderlineValues.Single), Underlined(UnderlineValues.Double), false);
    }

    [Test]
    public async Task VerticalAlignment()
    {
        await Check(Aligned(VerticalPositionValues.Superscript), Aligned(VerticalPositionValues.Superscript), true);
        await Check(Aligned(VerticalPositionValues.Superscript), Aligned(VerticalPositionValues.Subscript), false);
    }

    // What the renderers build: each property assigned through its typed setter, in whatever order
    // the nesting reached the run. The setters place by the rPr sequence, so the order is not seen.
    [Test]
    public Task TypedPropertiesAssignedInDifferentOrder()
    {
        Action<RunProperties> bold = _ => _.Bold = new();
        Action<RunProperties> underline = _ => _.Underline = new()
        {
            Val = UnderlineValues.Single
        };
        Action<RunProperties> style = _ => _.RunStyle = new()
        {
            Val = "Lead"
        };

        return Check(Assigned(bold, underline, style), Assigned(style, underline, bold), true);
    }

    static RunProperties Assigned(params Action<RunProperties>[] steps)
    {
        var properties = new RunProperties();
        foreach (var step in steps)
        {
            step(properties);
        }

        return properties;
    }

    [Test]
    public async Task ExtendedAttribute()
    {
        await Check(Extended("a", "1"), Extended("a", "1"), true);
        await Check(Extended("a", "1"), Extended("a", "2"), false);
        await Check(Extended("a", "1"), Extended("b", "1"), false);
        await Check(Extended("a", "1"), new(new Bold()), false);
    }

    [Test]
    public async Task ExtendedAttributeNamespaceAndPrefix()
    {
        await Check(Extended("a", "1"), Extended("a", "1", "w14", word14Namespace), false);
        await Check(Extended("a", "1", "w14", word14Namespace), Extended("a", "1", "w14", word14Namespace), true);
        await Check(Extended("a", "1", "w14", word14Namespace), Extended("a", "1", "x", word14Namespace), false);
    }

    // Extended attributes are written in the order they were added, so two sets holding the same
    // attributes in a different order are different xml, and were never merged.
    [Test]
    public Task ExtendedAttributeOrder()
    {
        var left = new Bold();
        left.SetAttribute(new("", "a", "", "1"));
        left.SetAttribute(new("", "b", "", "2"));
        var right = new Bold();
        right.SetAttribute(new("", "b", "", "2"));
        right.SetAttribute(new("", "a", "", "1"));

        return Check(new(left), new(right), false);
    }

    [Test]
    public async Task MarkupCompatibilityAttribute()
    {
        await Check(Ignorable("w14"), Ignorable("w14"), true);
        await Check(Ignorable("w14"), Ignorable("w15"), false);
        await Check(Ignorable("w14"), new(new Bold()), false);
        await Check(new(new Bold()), Ignorable("w14"), false);
    }

    [Test]
    public async Task NamespaceDeclaration()
    {
        await Check(Declared("w14", word14Namespace), Declared("w14", word14Namespace), true);
        await Check(Declared("w14", word14Namespace), Declared("x", word14Namespace), false);
        await Check(Declared("w14", word14Namespace), new(new Bold()), false);
        await Check(new(new Bold()), Declared("w14", word14Namespace), false);
    }

    [Test]
    public async Task NestedDifference()
    {
        await Check(Changed("Alex", "1"), Changed("Alex", "1"), true);
        await Check(Changed("Alex", "1"), Changed("Sam", "1"), false);
        await Check(Changed("Alex", "1"), Changed("Alex", "2"), false);

        var deeper = Changed("Alex", "1");
        deeper.RunPropertiesChange!.PreviousRunProperties!.Append(new Italic());
        await Check(Changed("Alex", "1"), deeper, false);
    }

    [Test]
    public async Task UnknownElement()
    {
        await Check(Unknown("x", "x", "text"), Unknown("x", "x", "text"), true);
        await Check(Unknown("x", "x", "text"), Unknown("x", "x", "other"), false);
        await Check(Unknown("x", "x", "text"), Unknown("x", "y", "text"), false);
        await Check(Unknown("x", "x", "text"), Unknown("v", "x", "text"), false);
        await Check(Unknown("x", "x", "text"), new(new Bold()), false);
    }

    // Read back from xml, which is how a run reaches the renderer from a clone of stored markup.
    [Test]
    public async Task ParsedFromXml()
    {
        var bold = $"""<w:rPr xmlns:w="{wordNamespace}"><w:rStyle w:val="Lead"/><w:b/></w:rPr>""";
        var italic = $"""<w:rPr xmlns:w="{wordNamespace}"><w:rStyle w:val="Lead"/><w:i/></w:rPr>""";
        var unknown = $"""<w:rPr xmlns:w="{wordNamespace}"><w:b/><w:other w:val="1">text</w:other></w:rPr>""";
        var unknownOther = $"""<w:rPr xmlns:w="{wordNamespace}"><w:b/><w:other w:val="2">text</w:other></w:rPr>""";

        await Check(new(bold), new(bold), true);
        await Check(new(bold), new(italic), false);
        await Check(new(unknown), new(unknown), true);
        await Check(new(unknown), new(unknownOther), false);
    }

    // An rPr built from a string hands that string back as its OuterXml until something reads into
    // it, so the string comparison told it apart from an identical one built in code. It stays
    // apart here, by the namespace declaration it was read with.
    [Test]
    public Task ParsedAgainstBuilt()
    {
        var parsed = new RunProperties($"""<w:rPr xmlns:w="{wordNamespace}"><w:rStyle w:val="Lead"/><w:b/></w:rPr>""");
        var built = new RunProperties(
            new RunStyle
            {
                Val = "Lead"
            },
            new Bold());

        return Check(parsed, built, false);
    }

    static async Task Check(RunProperties? left, RunProperties? right, bool expected)
    {
        await Assert.That(left?.OuterXml == right?.OuterXml).IsEqualTo(expected);
        await Assert.That(RunPropertiesComparer.Same(left, right)).IsEqualTo(expected);
        await Assert.That(RunPropertiesComparer.Same(right, left)).IsEqualTo(expected);
    }

    static RunProperties Styled(string id) =>
        new(
            new RunStyle
            {
                Val = id
            });

    static RunProperties Fonts(string ascii, string highAnsi) =>
        new(
            new RunFonts
            {
                Ascii = ascii,
                HighAnsi = highAnsi
            });

    static RunProperties Colored(string value) =>
        new(
            new Color
            {
                Val = value
            });

    static RunProperties Highlighted(HighlightColorValues value) =>
        new(
            new Highlight
            {
                Val = value
            });

    static RunProperties Underlined(UnderlineValues value) =>
        new(
            new Underline
            {
                Val = value
            });

    static RunProperties Aligned(VerticalPositionValues value) =>
        new(
            new VerticalTextAlignment
            {
                Val = value
            });

    static RunProperties Extended(string name, string value, string prefix = "", string namespaceUri = "")
    {
        var bold = new Bold();
        bold.SetAttribute(new(prefix, name, namespaceUri, value));
        return new(bold);
    }

    static RunProperties Ignorable(string value) =>
        new(
            new Bold
            {
                MCAttributes = new()
                {
                    Ignorable = value
                }
            });

    static RunProperties Declared(string prefix, string uri)
    {
        var bold = new Bold();
        bold.AddNamespaceDeclaration(prefix, uri);
        return new(bold);
    }

    static RunProperties Changed(string author, string id) =>
        new(
            new Bold(),
            new RunPropertiesChange(
                new PreviousRunProperties(
                    new Bold()))
            {
                Author = author,
                Id = id
            });

    static RunProperties Unknown(string prefix, string name, string text)
    {
        var properties = new RunProperties($"""<w:rPr xmlns:w="{wordNamespace}"><w:b/><{prefix}:{name} xmlns:{prefix}="urn:unknown">{text}</{prefix}:{name}></w:rPr>""");

        // Reading a child is what turns the string into elements.
        _ = properties.FirstChild;
        return properties;
    }
}
