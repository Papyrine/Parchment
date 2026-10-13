// ReSharper disable PartialTypeWithSinglePart
// ReSharper disable InconsistentNaming
#pragma warning disable CA1051, CA2211, IDE1006

/// <summary>
/// A name that more than one member answers to still resolves to one member: the one the source
/// generator emits an accessor for. A <c>new</c> member hides the base one whatever its type, and
/// of two members whose names differ only by case the first declared is the member.
/// </summary>
public partial class AmbiguousMemberTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required Leaf Leaf { get; init; }
        public required Cased Cased { get; init; }

        public string Title { get; init; } = "upper";
        public int title { get; init; } = 7;
    }

    public class Root
    {
        public decimal Total { get; init; } = 1m;
        public decimal Tax { get; init; } = 2m;
        public decimal Rate = 3m;
        public static decimal Shared => 4m;
        public static decimal Limit = 5m;
        public string Kept { get; init; } = "kept";
    }

    public class Middle :
        Root
    {
        public new string Tax { get; init; } = "middle tax";
    }

    public class Leaf :
        Middle
    {
        public new string Total { get; init; } = "leaf total";
        public new string Rate = "leaf rate";
        public new static string Shared => "leaf shared";
        public new static string Limit = "leaf limit";
    }

    public class Cased
    {
        public string Name { get; init; } = "first";
        public int name { get; init; } = 1;

        public string Code = "first";
        public int CODE = 2;

        public string Total { get; init; } = "first";
        public int total = 3;

        public string count = "first";
        public int Count { get; init; } = 4;

        public static string Shared => "first";
        public static int SHARED => 5;

        // Computed, so nothing in metadata places it against the field: the property is taken.
        public string Label => Code;
        public int label = 6;
    }

    static Doc BuildModel() =>
        new()
        {
            Leaf = new(),
            Cased = new()
        };

    [Test]
    [Arguments(typeof(Leaf), "Total", typeof(string))]
    [Arguments(typeof(Leaf), "total", typeof(string))]
    [Arguments(typeof(Leaf), "Tax", typeof(string))]
    [Arguments(typeof(Leaf), "Rate", typeof(string))]
    [Arguments(typeof(Leaf), "Shared", typeof(string))]
    [Arguments(typeof(Leaf), "Limit", typeof(string))]
    [Arguments(typeof(Leaf), "Kept", typeof(string))]
    [Arguments(typeof(Middle), "Total", typeof(decimal))]
    [Arguments(typeof(Middle), "Tax", typeof(string))]
    [Arguments(typeof(Middle), "Shared", typeof(decimal))]
    [Arguments(typeof(Middle), "limit", typeof(decimal))]
    [Arguments(typeof(Root), "Tax", typeof(decimal))]
    [Arguments(typeof(Cased), "Name", typeof(string))]
    [Arguments(typeof(Cased), "name", typeof(string))]
    [Arguments(typeof(Cased), "NAME", typeof(string))]
    [Arguments(typeof(Cased), "Code", typeof(string))]
    [Arguments(typeof(Cased), "CODE", typeof(string))]
    [Arguments(typeof(Cased), "Total", typeof(string))]
    [Arguments(typeof(Cased), "total", typeof(string))]
    [Arguments(typeof(Cased), "count", typeof(string))]
    [Arguments(typeof(Cased), "Count", typeof(string))]
    [Arguments(typeof(Cased), "Shared", typeof(string))]
    [Arguments(typeof(Cased), "SHARED", typeof(string))]
    [Arguments(typeof(Cased), "label", typeof(string))]
    [Arguments(typeof(Doc), "title", typeof(string))]
    public async Task ResolvesToTheMemberTheGeneratorEmits(Type type, string name, Type expected)
    {
        await Assert.That(ModelValidator.TryResolveMember(type, name, out var memberType)).IsTrue();
        await Assert.That(memberType).IsEqualTo(expected);
    }

    [Test]
    [Arguments(typeof(Leaf), "Missing")]
    [Arguments(typeof(Cased), "Nam")]
    public async Task UnknownNameStillDoesNotResolve(Type type, string name) =>
        await Assert.That(ModelValidator.TryResolveMember(type, name, out _)).IsFalse();

    // The spelling is the one the accessor is registered under, so the value is the resolved
    // member's in both flows.
    [Test]
    [Arguments("{{ Leaf.Total }}", "leaf total")]
    [Arguments("{{ Leaf.Total.size }}", "10")]
    [Arguments("{{ Leaf.Tax }}", "middle tax")]
    [Arguments("{{ Leaf.Rate }}", "leaf rate")]
    [Arguments("{{ Leaf.Shared }}", "leaf shared")]
    [Arguments("{{ Leaf.Limit }}", "leaf limit")]
    [Arguments("{{ Leaf.Kept }}", "kept")]
    [Arguments("{{ Title }}", "upper")]
    [Arguments("{{ Cased.Name }}", "first")]
    [Arguments("{{ Cased.Name.size }}", "5")]
    [Arguments("{{ Cased.Code }}", "first")]
    [Arguments("{{ Cased.Total }}", "first")]
    [Arguments("{{ Cased.count }}", "first")]
    [Arguments("{{ Cased.Shared }}", "first")]
    [Arguments("{{ Cased.Label }}", "first")]
    public async Task RegistersAndRenders(string token, string expected)
    {
        await Assert.That(await RenderDocx($"[{token}]")).IsEqualTo($"[{expected}]");
        await Assert.That(await RenderMarkdown($"[{token}]")).IsEqualTo($"[{expected}]");
    }

    // The other spelling names the same member, so what follows it is checked against that
    // member's type and not against its namesake's.
    [Test]
    [Arguments("{{ title.size }}")]
    [Arguments("{{ Cased.name.size }}")]
    [Arguments("{{ Cased.CODE.size }}")]
    [Arguments("{{ Cased.total.size }}")]
    [Arguments("{{ Cased.Count.size }}")]
    [Arguments("{{ Cased.SHARED.size }}")]
    [Arguments("{{ leaf.total.size }}")]
    public async Task OtherSpellingRegisters(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        new TemplateStore().RegisterDocxTemplate<Doc>(template);

        using var styleSource = DocxTemplateBuilder.Build();
        new TemplateStore().RegisterMarkdownTemplate<Doc>(content, styleSource);

        await Assert.That(SharedFluid.IsModelRegistered(typeof(Doc))).IsTrue();
    }

    [Test]
    [Arguments("Leaf.Total.Nope", "Nope")]
    [Arguments("Leaf.Tax.Nope", "Nope")]
    [Arguments("Cased.name.Nope", "Nope")]
    [Arguments("Cased.Nam", "Nam")]
    public async Task MissingMemberIsARegistrationError(string path, string missing)
    {
        using var template = DocxTemplateBuilder.Build($"{{{{ {path} }}}}");
        var docx = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Doc>(template));
        await Assert.That(docx.Message).Contains($"'{missing}'");

        using var styleSource = DocxTemplateBuilder.Build();
        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Doc>($"{{{{ {path} }}}}", styleSource));
        await Assert.That(markdown.Message).Contains($"'{missing}'");
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
