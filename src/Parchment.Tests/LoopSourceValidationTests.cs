// ReSharper disable PartialTypeWithSinglePart
// Registration verdicts for what a {% for %} tag names, in both flows. The same templates run
// through the source generator in Parchment.SourceGenerator.Tests/LoopSourceValidationTests.cs, and
// the two files have to agree case for case.
public partial class LoopSourceValidationTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public required string Name { get; init; }
        public required IReadOnlyList<string> Lines { get; init; }
        public required char[] Letters { get; init; }
        public required IEnumerable<char> Chars { get; init; }
        public required IEnumerable Bag { get; init; }
        public Memory<int> Buffer { get; init; }
        public int Max { get; init; }
        public int Skip { get; init; }
    }

    static Model Sample() =>
        new()
        {
            Name = "Doc",
            Lines = ["one", "two", "three"],
            Letters = ['a', 'b'],
            Chars = ['c', 'd'],
            Bag = new ArrayList
            {
                "x"
            },
            Max = 2,
            Skip = 1
        };

    // A string is an IEnumerable<char>, but liquid runs the loop once with the whole string, so
    // it is rejected like any other member that is not a collection. The other two are not an
    // IEnumerable<T> at all.
    [Test]
    [Arguments("Name")]
    [Arguments("Bag")]
    [Arguments("Buffer")]
    public async Task SourceThatIsNotACollection_FailsRegistration(string source)
    {
        var docx = Assert.Throws<ParchmentRegistrationException>(() => RegisterDocx($"{{% for item in {source} %}}"));
        await Assert.That(docx.Message).Contains("does not resolve to an enumerable");

        var markdown = Assert.Throws<ParchmentRegistrationException>(() => RegisterMarkdown($"{{% for item in {source} %}}"));
        await Assert.That(markdown.Message).Contains("does not resolve to an enumerable");
    }

    // A sequence of char is a collection like any other: each flow renders one item per character.
    [Test]
    [Arguments("Letters", "a,b")]
    [Arguments("Chars", "c,d")]
    public async Task CharSequence_IteratesItsCharacters(string source, string expected)
    {
        await Assert.That(await Render(RegisterDocx($"{{% for item in {source} %}}"))).IsEqualTo(expected);
        await Assert.That(await Render(RegisterMarkdown($"{{% for item in {source} %}}"))).IsEqualTo(expected);
    }

    // The markdown flow hands the tag to liquid, where a name that does not resolve in a modifier
    // is nil: `limit: Missing` renders nothing and `offset: Missing` renders everything.
    [Test]
    [Arguments("Lines limit: Missing")]
    [Arguments("Lines offset: Missing")]
    // The CLR member is not what liquid binds; `size` is.
    [Arguments("Lines limit: Lines.Count")]
    // Evaluated before the loop variable is bound.
    [Arguments("Lines limit: item")]
    [Arguments("(1..Missing)")]
    [Arguments("(1..5) limit: Missing")]
    [Arguments("(1..5) offset: Missing")]
    public async Task UnknownMemberInModifierOrBound_FailsRegistration(string loop)
    {
        await Assert.That(() => RegisterDocx($"{{% for item in {loop} %}}"))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(() => RegisterMarkdown($"{{% for item in {loop} %}}"))
            .Throws<ParchmentRegistrationException>();
    }

    [Test]
    [Arguments("Lines limit: Max offset: Skip", "two,three")]
    [Arguments("Lines limit:2 offset:1 reversed", "three,two")]
    [Arguments("Lines reversed", "three,two,one")]
    [Arguments("Lines limit: Lines.size", "one,two,three")]
    [Arguments("(1..Max) offset: Skip", "2")]
    public async Task KnownMemberInModifierOrBound_Renders(string loop, string expected)
    {
        await Assert.That(await Render(RegisterDocx($"{{% for item in {loop} %}}"))).IsEqualTo(expected);
        await Assert.That(await Render(RegisterMarkdown($"{{% for item in {loop} %}}"))).IsEqualTo(expected);
    }

    // `continue` is liquid's keyword for resuming after the previous loop over the same source,
    // not a model member. Only the markdown flow has the loop state it needs.
    [Test]
    public async Task OffsetContinue_IsAcceptedInMarkdown()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>(
            "{% for item in Lines limit:1 %}{{ item }}\n\n{% endfor %}{% for item in Lines offset: continue %}{{ item }}\n\n{% endfor %}");
        await Assert.That(await Render(store)).IsEqualTo("one,two,three");
    }

    static TemplateStore RegisterDocx(string openTag)
    {
        using var template = DocxTemplateBuilder.Build(
            $$$"""
               {{{openTag}}}

               {{ item }}

               {% endfor %}
               """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);
        return store;
    }

    static TemplateStore RegisterMarkdown(string openTag)
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>($"{openTag}\n{{{{ item }}}}\n\n{{% endfor %}}");
        return store;
    }

    static async Task<string> Render(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(Sample(), stream);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, false);
        var texts = doc.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0);
        return string.Join(',', texts);
    }
}
