// ReSharper disable PartialTypeWithSinglePart

/// <summary>
/// Registration has to agree with the render about the members of a system type. The source
/// generator emits accessors for the model's own types only, so <c>When.Year</c> on a
/// <see cref="DateTime"/> binds nothing, and neither does a <c>model.</c> or type-name prefix.
/// <c>SystemTypeMemberGeneratorTests</c> runs the same paths through the generator.
/// </summary>
public partial class SystemTypeMemberTests
{
    [ParchmentBindable]
    public partial class Invoice
    {
        public string Name { get; init; } = "Alice";
        public DateTime When { get; init; } = new(2024, 3, 5, 10, 30, 0);
        public DateTimeOffset Stamp { get; init; } = new(2024, 3, 5, 10, 30, 0, TimeSpan.FromHours(2));
        public DateOnly Date { get; init; } = new(2024, 3, 5);
        public TimeOnly Time { get; init; } = new(10, 30);
        public TimeSpan Span { get; init; } = TimeSpan.FromDays(2);
        public Guid Id { get; init; } = new("0f8fad5b-d9cb-469f-a165-70867728950e");
        public Uri Link { get; init; } = new("https://example.com/path");
        public int Count { get; init; } = 3;
        public decimal Amount { get; init; } = 12.50m;
        public int? Maybe { get; init; } = 5;
        public Meta? Meta { get; init; } = new Meta
        {
            Name = "meta"
        };
        public (int, string) Pair { get; init; } = (7, "seven");
        public Tuple<int, string> RefPair { get; init; } = Tuple.Create(8, "eight");
        public KeyValuePair<string, int> Kv { get; init; } = new("key", 9);
        public DayOfWeek Day { get; init; } = DayOfWeek.Friday;
        public List<DateTime> Dates { get; init; } = [new(2024, 3, 5)];
        public Dictionary<string, Line> Lookup { get; init; } = new()
        {
            ["key"] = new()
            {
                Name = "looked up"
            }
        };
        public List<Line> Lines { get; init; } =
        [
            new()
            {
                Name = "one"
            }
        ];
        public IShape Shape { get; init; } = new Circle();
        public Special Special { get; init; } = new();
    }

    public struct Meta
    {
        public string Name { get; init; }
    }

    public class Line
    {
        public required string Name { get; init; }
    }

    public interface IShape
    {
        string Kind { get; }
    }

    public class Circle :
        IShape
    {
        public string Kind => "circle";
    }

    public class LineBase
    {
        public string BaseName { get; init; } = "base";
    }

    public class Special :
        LineBase;

    // Every one of these names a real CLR member. The first test pins that Fluid binds none of
    // them, which is the only reason the second may reject them.
    [Test]
    [Arguments("When.Year")]
    [Arguments("When.Date")]
    [Arguments("Stamp.Offset")]
    [Arguments("Date.Year")]
    [Arguments("Time.Hour")]
    [Arguments("Span.TotalDays")]
    [Arguments("Id.Version")]
    [Arguments("Link.Host")]
    [Arguments("Count.MaxValue")]
    [Arguments("Amount.Scale")]
    [Arguments("Maybe.Value")]
    [Arguments("Maybe.HasValue")]
    [Arguments("Meta.Value.Name")]
    [Arguments("Pair.Item1")]
    [Arguments("RefPair.Item1")]
    [Arguments("Day.Friday")]
    [Arguments("Dates.first.Year")]
    [Arguments("Lines.size.MaxValue")]
    public async Task SystemTypeMemberRendersBlankInFluid(string path) =>
        await Assert.That(await RenderFluid($"[{{{{ {path} }}}}]")).IsEqualTo("[]");

    [Test]
    [Arguments("When.Year", "Year", "DateTime")]
    [Arguments("When.Date", "Date", "DateTime")]
    [Arguments("Stamp.Offset", "Offset", "DateTimeOffset")]
    [Arguments("Date.Year", "Year", "DateOnly")]
    [Arguments("Time.Hour", "Hour", "TimeOnly")]
    [Arguments("Span.TotalDays", "TotalDays", "TimeSpan")]
    [Arguments("Id.Version", "Version", "Guid")]
    [Arguments("Link.Host", "Host", "Uri")]
    [Arguments("Count.MaxValue", "MaxValue", "Int32")]
    [Arguments("Amount.Scale", "Scale", "Decimal")]
    [Arguments("Maybe.Value", "Value", "Nullable`1")]
    [Arguments("Maybe.HasValue", "HasValue", "Nullable`1")]
    [Arguments("Meta.Value.Name", "Value", "Nullable`1")]
    [Arguments("Pair.Item1", "Item1", "ValueTuple`2")]
    [Arguments("RefPair.Item1", "Item1", "Tuple`2")]
    [Arguments("Day.Friday", "Friday", "DayOfWeek")]
    [Arguments("Dates.first.Year", "Year", "DateTime")]
    [Arguments("Lines.size.MaxValue", "MaxValue", "Int32")]
    public async Task SystemTypeMemberFailsRegistration(string path, string member, string type)
    {
        var expected = $"Member '{member}' of '{type}' is not available to a template.";

        await Assert.That(DocxFailure($"{{{{ {path} }}}}")).Contains(expected);
        await Assert.That(MarkdownFailure($"{{{{ {path} }}}}")).Contains(expected);
    }

    // The value of a nullable user struct arrives boxed as the struct, but the generator never
    // reaches the struct through Nullable<T> and emits no accessors for it. Rejected before this
    // rule existed, by both the runtime and the generator, and still rejected.
    [Test]
    public async Task MemberOfANullableUserStructFailsRegistration()
    {
        await Assert.That(await RenderFluid("[{{ Meta.Name }}]")).IsEqualTo("[]");
        await Assert.That(DocxFailure("{{ Meta.Name }}")).Contains("'Name'");
        await Assert.That(MarkdownFailure("{{ Meta.Name }}")).Contains("'Name'");
    }

    [Test]
    public async Task SystemTypeMemberOfALoopVariableFailsRegistration()
    {
        var docx = DocxFailure(
            """
            {% for date in Dates %}

            {{ date.Year }}

            {% endfor %}
            """);
        await Assert.That(docx).Contains("Member 'Year' of 'DateTime' is not available to a template.");

        var markdown = MarkdownFailure("{% for date in Dates %}{{ date.Year }}{% endfor %}");
        await Assert.That(markdown).Contains("Member 'Year' of 'DateTime' is not available to a template.");
    }

    [Test]
    public async Task SystemTypeMemberInAConditionFailsRegistration()
    {
        var docx = DocxFailure(
            """
            {% if Maybe.HasValue %}

            some

            {% endif %}
            """);
        await Assert.That(docx).Contains("'HasValue'");

        var markdown = MarkdownFailure("{% if Maybe.HasValue %}some{% endif %}");
        await Assert.That(markdown).Contains("'HasValue'");
    }

    [Test]
    public async Task LoopSourceThroughASystemTypeFailsRegistration()
    {
        var docx = DocxFailure(
            """
            {% for item in RefPair.Item2 %}

            {{ item }}

            {% endfor %}
            """);
        await Assert.That(docx).Contains("RefPair.Item2");

        var markdown = MarkdownFailure("{% for item in RefPair.Item2 %}{{ item }}{% endfor %}");
        await Assert.That(markdown).Contains("'Item2'");
    }

    // What still has to bind: the value itself, liquid's own members, the two members the
    // generator surfaces on a KeyValuePair, and the members of a model's own types however they
    // are reached.
    [Test]
    [Arguments("{{ Maybe }}", "5")]
    [Arguments("{{ Count }}", "3")]
    [Arguments("{{ Day }}", "Friday")]
    [Arguments("{{ Name.size }}", "5")]
    [Arguments("{{ Dates.size }}", "1")]
    [Arguments("{{ Kv.Key }}", "key")]
    [Arguments("{{ Kv.Value }}", "9")]
    [Arguments("{{ Lines.first.Name }}", "one")]
    [Arguments("{{ Shape.Kind }}", "circle")]
    [Arguments("{{ Special.BaseName }}", "base")]
    public async Task BoundMemberRegistersAndRenders(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    [Test]
    public async Task MemberOfAUserTypeThroughALoopRegistersAndRenders()
    {
        var docx = await RenderDocx(
            """
            {% for line in Lines %}

            {{ line.Name }}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("one");

        var markdown = await RenderMarkdown("{% for line in Lines %}{{ line.Name }}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("one");
    }

    // Registration only: Key and Value are the members the generator surfaces for a dictionary
    // loop, so they have to keep validating whatever the render makes of them.
    [Test]
    public void KeyAndValueOfADictionaryLoopRegister()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for pair in Lookup %}

            {{ pair.Key }} {{ pair.Value.Name }}

            {% endfor %}
            """);
        new TemplateStore().RegisterDocxTemplate<Invoice>(template);

        using var styleSource = DocxTemplateBuilder.Build();
        new TemplateStore().RegisterMarkdownTemplate<Invoice>(
            "{% for pair in Lookup %}{{ pair.Key }} {{ pair.Value.Name }}{% endfor %}",
            styleSource);
    }

    // Nothing puts the model under a name of its own: it is the scope the root identifiers are
    // read from. Pinned against Fluid first, as above.
    [Test]
    [Arguments("model.Name")]
    [Arguments("Model.Name")]
    [Arguments("Invoice.Name")]
    [Arguments("invoice.Name")]
    [Arguments("model")]
    [Arguments("Invoice")]
    public async Task ModelPrefixRendersBlankInFluid(string path) =>
        await Assert.That(await RenderFluid($"[{{{{ {path} }}}}]")).IsEqualTo("[]");

    [Test]
    [Arguments("model.Name", "model")]
    [Arguments("Model.Name", "Model")]
    [Arguments("Invoice.Name", "Invoice")]
    [Arguments("invoice.Name", "invoice")]
    [Arguments("model", "model")]
    [Arguments("Invoice", "Invoice")]
    public async Task ModelPrefixFailsRegistration(string path, string root)
    {
        var expected = $"Identifier '{root}' is not a member of 'Invoice'";

        await Assert.That(DocxFailure($"{{{{ {path} }}}}")).Contains(expected);
        await Assert.That(MarkdownFailure($"{{{{ {path} }}}}")).Contains(expected);
    }

    [Test]
    public async Task ModelPrefixOnALoopSourceFailsRegistration()
    {
        var docx = DocxFailure(
            """
            {% for line in model.Lines %}

            {{ line.Name }}

            {% endfor %}
            """);
        await Assert.That(docx).Contains("model.Lines");

        var markdown = MarkdownFailure("{% for line in model.Lines %}{{ line.Name }}{% endfor %}");
        await Assert.That(markdown).Contains("Identifier 'model' is not a member of 'Invoice'");
    }

    // A loop variable is found before the model is asked, whatever it is called.
    [Test]
    public async Task LoopVariableNamedModelRegistersAndRenders()
    {
        var docx = await RenderDocx(
            """
            {% for model in Lines %}

            {{ model.Name }}

            {% endfor %}
            """);
        await Assert.That(docx).IsEqualTo("one");

        var markdown = await RenderMarkdown("{% for model in Lines %}{{ model.Name }}{% endfor %}");
        await Assert.That(markdown).IsEqualTo("one");
    }

    static string DocxFailure(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        return Assert.Throws<ParchmentRegistrationException>(
                () => new TemplateStore().RegisterDocxTemplate<Invoice>(template))
            .Message;
    }

    static string MarkdownFailure(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        return Assert.Throws<ParchmentRegistrationException>(
                () => new TemplateStore().RegisterMarkdownTemplate<Invoice>(markdown, styleSource))
            .Message;
    }

    static async Task<string> RenderFluid(string source)
    {
        SharedFluid.Parser.TryParse(source, out var template, out _);
        var context = new TemplateContext(new Invoice(), SharedFluid.Options, allowModelMembers: true);
        return await template!.RenderAsync(context);
    }

    static async Task<string> RenderDocx(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Invoice>(template);
        return await RenderText(store);
    }

    static async Task<string> RenderMarkdown(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Invoice>(markdown, styleSource);
        return await RenderText(store);
    }

    static async Task<string> RenderText(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(new Invoice(), stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
