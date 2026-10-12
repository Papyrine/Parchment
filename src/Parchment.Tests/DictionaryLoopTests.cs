// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// Every validator binds the variable of a loop over a dictionary to a <c>KeyValuePair</c>, so
/// <c>kv.Key</c> and <c>kv.Value</c> register. Fluid hands the loop a two-element array per entry
/// instead, and the render has to answer to the same two names.
/// </summary>
public partial class DictionaryLoopTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required Dictionary<string, string> Lookup { get; init; }
        public required IDictionary<string, string> Interface { get; init; }
        public required IReadOnlyDictionary<string, string> ReadOnly { get; init; }
        public required SortedDictionary<string, string> Sorted { get; init; }
        public required Dictionary<string, Line> Lines { get; init; }
        public required Holder Holder { get; init; }
        public required List<Dictionary<string, string>> Listed { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
    }

    public class Holder
    {
        public required Dictionary<string, string> Lookup { get; init; }
    }

    // Keys are added out of order: a dictionary iterates in its own order, a sorted one by key.
    static Dictionary<string, string> Pairs() =>
        new()
        {
            ["b"] = "two",
            ["a"] = "one",
            ["c"] = "three"
        };

    static Doc BuildModel() =>
        new()
        {
            Lookup = Pairs(),
            Interface = Pairs(),
            ReadOnly = Pairs(),
            Sorted = new(Pairs()),
            Lines = new()
            {
                ["x"] = new()
                {
                    Name = "first"
                },
                ["y"] = new()
                {
                    Name = "second"
                }
            },
            Holder = new()
            {
                Lookup = Pairs()
            },
            Listed = [Pairs()]
        };

    [Test]
    [Arguments("Lookup", "", "{{ kv.Key }}={{ kv.Value }};", "b=two;a=one;c=three;")]
    [Arguments("Interface", "", "{{ kv.Key }}={{ kv.Value }};", "b=two;a=one;c=three;")]
    [Arguments("ReadOnly", "", "{{ kv.Key }}={{ kv.Value }};", "b=two;a=one;c=three;")]
    [Arguments("Sorted", "", "{{ kv.Key }}={{ kv.Value }};", "a=one;b=two;c=three;")]
    [Arguments("Lines", "", "{{ kv.Key }}={{ kv.Value.Name }};", "x=first;y=second;")]
    [Arguments("Holder.Lookup", "", "{{ kv.Key }}={{ kv.Value }};", "b=two;a=one;c=three;")]
    [Arguments("Lookup", "", "{{ kv.key }}={{ kv.VALUE }};", "b=two;a=one;c=three;")]
    [Arguments("Lookup", "", "{{ kv['Key'] }}={{ kv['Value'] }};", "b=two;a=one;c=three;")]
    [Arguments("Lookup", "", "{{ kv[0] }}={{ kv[1] }};", "b=two;a=one;c=three;")]
    [Arguments("Lines", "", "{{ kv[0] }}={{ kv[1].Name }};", "x=first;y=second;")]
    [Arguments("Lookup", "", "{{ kv.Value | upcase }};", "TWO;ONE;THREE;")]
    [Arguments("Lookup", "", "{{ kv.Value.size }};", "3;3;5;")]
    [Arguments("Lookup", " limit:2", "{{ kv.Key }}={{ kv.Value }};", "b=two;a=one;")]
    [Arguments("Lookup", " offset:1", "{{ kv.Key }}={{ kv.Value }};", "a=one;c=three;")]
    [Arguments("Lookup", " reversed", "{{ kv.Key }}={{ kv.Value }};", "c=three;a=one;b=two;")]
    [Arguments("Sorted", " reversed limit:2", "{{ kv.Key }}={{ kv.Value }};", "b=two;a=one;")]
    public async Task EntryRegistersAndRenders(string source, string modifiers, string body, string expected)
    {
        var docx = await RenderDocx(
            $$"""
              {% for kv in {{source}}{{modifiers}} %}

              {{body}}

              {% endfor %}
              """);
        await Assert.That(docx).IsEqualTo(expected);

        var markdown = await RenderMarkdown($"{{% for kv in {source}{modifiers} %}}{body}{{% endfor %}}");
        await Assert.That(markdown).IsEqualTo(expected);
    }

    [Test]
    public async Task EntryDrivesAConditional()
    {
        var docx = await RenderDocx(
            """
            {% for kv in Lookup %}

            {% if kv.Key == "a" %}

            {{ kv.Value }}

            {% endif %}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("one");

        var markdown = await RenderMarkdown("{% for kv in Lookup %}{% if kv.Key == \"a\" %}{{ kv.Value }}{% endif %}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("one");
    }

    // Fluid converts the items of a list as it wraps the list, so a dictionary inside one is
    // reached by a different route from a dictionary read off a member.
    [Test]
    public async Task DictionaryInsideAList()
    {
        var docx = await RenderDocx(
            """
            {% for lookup in Listed %}

            {% for kv in lookup %}

            {{ kv.Key }}={{ kv.Value }};

            {% endfor %}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("b=two;a=one;c=three;");

        var markdown = await RenderMarkdown("{% for lookup in Listed %}{% for kv in lookup %}{{ kv.Key }}={{ kv.Value }};{% endfor %}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("b=two;a=one;c=three;");
    }

    // What a dictionary answered before it was wrapped, it still answers.
    [Test]
    [Arguments("{{ Lookup.size }}", "3")]
    [Arguments("{{ Lines.size }}", "2")]
    public async Task DictionaryKeepsItsLiquidMembers(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    [Test]
    [Arguments("[{{ Lookup.a }}]", "[one]")]
    [Arguments("[{{ Lookup['c'] }}]", "[three]")]
    [Arguments("[{{ Lines.y.Name }}]", "[second]")]
    [Arguments("[{{ Lookup }}]", "[]")]
    [Arguments("{% if Lookup %}yes{% endif %}", "yes")]
    [Arguments("{% for kv in Lookup %}{{ kv.first }}={{ kv.last }}/{{ kv.size }};{% endfor %}", "b=two/2;a=one/2;c=three/2;")]
    [Arguments("{% for kv in Lookup %}{% for part in kv %}{{ part }},{% endfor %}{% endfor %}", "b,two,a,one,c,three,")]
    [Arguments("{% for kv in Lookup %}{{ kv }};{% endfor %}", "btwo;aone;cthree;")]
    [Arguments("{% for kv in Lookup %}{{ kv.Nope }};{% endfor %}", ";;;")]
    public async Task FluidBehaviourIsKept(string source, string expected)
    {
        SharedFluid.Parser.TryParse(source, out var template, out _);
        var context = new TemplateContext(BuildModel(), SharedFluid.Options, allowModelMembers: true);
        await Assert.That(await template!.RenderAsync(context)).IsEqualTo(expected);
    }

    // `first` and `last` are the array's names for the two halves, not the pair's, and the loop
    // variable is typed as the pair. They stay rejected, so nothing registers that the names the
    // validators know cannot say.
    [Test]
    [Arguments("first")]
    [Arguments("last")]
    [Arguments("Nope")]
    public async Task OtherEntryMemberFailsRegistration(string member)
    {
        using var template = DocxTemplateBuilder.Build(
            $$$"""
               {% for kv in Lookup %}

               {{ kv.{{{member}}} }}

               {% endfor %}
               """);
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains($"'{member}'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>($"{{% for kv in Lookup %}}{{{{ kv.{member} }}}}{{% endfor %}}", styleSource));
        await Assert.That(markdown.Message).Contains($"'{member}'");
    }

    [Test]
    public async Task UnknownMemberThroughValueFailsRegistration()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>("{% for kv in Lines %}{{ kv.Value.Nope }}{% endfor %}", styleSource));
        await Assert.That(markdown.Message).Contains("Nope");
    }

    static async Task<string> RenderDocx(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return await RenderText(store);
    }

    static async Task<string> RenderMarkdown(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return await RenderText(store);
    }

    static async Task<string> RenderText(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(BuildModel(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
