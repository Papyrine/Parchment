// ReSharper disable PartialTypeWithSinglePart
public partial class ReferenceValidatorTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public required string Title { get; init; }
        public required Profile Profile { get; init; }
        public required IReadOnlyList<Item> Items { get; init; }
    }

    public class Profile
    {
        public required string DisplayName { get; init; }
    }

    public class Item
    {
        public required string Sku { get; init; }
    }

    [ParchmentBindable]
    public partial class SelfRef
    {
        public required string Name { get; init; }
        public SelfRef? Next { get; init; }
    }

    [Test]
    public async Task UnknownRootMember_FailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build("{{ Missing }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("Missing");
    }

    [Test]
    public async Task UnknownNestedMember_FailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build("{{ Profile.DoesNotExist }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("DoesNotExist");
    }

    [Test]
    public async Task DeepValidPath_RegistersSuccessfully()
    {
        using var template = DocxTemplateBuilder.Build("{{ Profile.DisplayName }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
    }

    [Test]
    public async Task IndexerWithStringLiteral_ValidatesAsDottedAccess()
    {
        // `Customer['DisplayName']` is Fluid-equivalent to `Customer.DisplayName` at render — both
        // resolve to the same member. Validation should treat them the same instead of skipping
        // the indexer segment and missing typos.
        using var template = DocxTemplateBuilder.Build("{{ Profile['DisplayName'] }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
    }

    [Test]
    public async Task IndexerWithStringLiteral_UnknownMember_FailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build("{{ Profile['NoSuchMember'] }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("NoSuchMember");
    }

    [Test]
    public async Task IndexerWithDoubleQuotedLiteral_UnknownMember_FailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build("{{ Profile[\"NoSuchMember\"] }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("NoSuchMember");
    }

    [Test]
    public async Task MixedDotAndIndexer_UnknownMember_FailsRegistration()
    {
        // Cross-syntax path — dotted then indexer, indexer then dotted. Both should validate
        // all segments.
        using var template = DocxTemplateBuilder.Build("{{ Profile.DisplayName['NoSuchMember'] }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("NoSuchMember");
    }

    [Test]
    public async Task LoopVariableShadowsRootScope()
    {
        // Inside the loop, `Title` refers to the loop element's member (Item.Sku is fine, but
        // the loop variable named `Title` would shadow the root if reused). Here we use a
        // different name to confirm the loop variable's element type is honored.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for it in Items %}

            {{ it.Sku }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
    }

    [Test]
    public async Task LoopVariableMember_UnknownMember_FailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for it in Items %}

            {{ it.NotARealField }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("NotARealField");
    }

    // Whitespace control is documented as transparent to validation, in a section covering docx as
    // well as markdown. Markdown goes through Fluid's parser, which strips the hyphens; the docx
    // flow scans tags itself and used to reject these as malformed block tags.
    [Test]
    public void WhitespaceControlOnBlockTagsIsAccepted()
    {
        foreach (var open in new[] { "{%-", "{%" })
        {
            foreach (var close in new[] { "-%}", "%}" })
            {
                using var template = DocxTemplateBuilder.Build(
                    open + " for it in Items " + close +
                    "\n\n{{ it.Sku }}\n\n" +
                    open + " endfor " + close);

                var store = new TemplateStore();
                store.RegisterDocxTemplate<Doc>(template);
            }
        }
    }

    // The markdown flow allows forloop inside a loop body. The docx flow iterates through its own
    // scope tree and never populates it, so this has to keep failing — but the message should name
    // forloop instead of reading as a missing model member and sending the reader after a property.
    [Test]
    public async Task ForLoopIdentifier_FailsRegistrationWithAnExplanation()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% for it in Items %}

            {{ forloop.index }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("{% for %}");
        await Assert.That(exception.Message).DoesNotContain("is not a member of");
    }

    [Test]
    public async Task NonEnumerableLoopSource_FailsRegistration()
    {
        // Profile is a POCO, not a collection. Looping over it should be rejected.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for p in Profile %}

            {{ p }}

            {% endfor %}
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("does not resolve to an enumerable");
    }

    [Test]
    public async Task SelfReferentialModel_DoesNotRecurseForever()
    {
        // SelfRef.Next : SelfRef — naive walks would loop. Validator's per-branch visited
        // discipline must terminate.
        using var template = DocxTemplateBuilder.Build("{{ Name }} {{ Next.Name }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<SelfRef>(template);
    }

    [ParchmentBindable]
    public partial class DocumentBase
    {
        public required string Title { get; init; }
    }

    [ParchmentBindable]
    public partial class Report : DocumentBase
    {
        public required string Body { get; init; }
    }

    public class ShadowBase
    {
        public string Title { get; init; } = "base";
    }

    [ParchmentBindable]
    public partial class ShadowingReport : ShadowBase
    {
        public new string Title { get; init; } = "derived";
    }

    public class Sealed
    {
        public string Name { get; set; } = "";

        public string Hidden { private get; set; } = "";

        public string WriteOnly
        {
            set => Name = value;
        }
    }

    [ParchmentBindable]
    public partial class Envelope
    {
        public Sealed Sealed { get; init; } = new();
    }

    // The generator emits no accessor for a property it cannot read, so one that registered would
    // render blank.
    [Test]
    [Arguments("Sealed.Hidden")]
    [Arguments("Sealed.WriteOnly")]
    public async Task UnreadableMember_FailsRegistration(string path)
    {
        using var template = DocxTemplateBuilder.Build($"{{{{ {path} }}}}");
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Envelope>(template));
        await Assert.That(exception.Message).Contains(path[(path.IndexOf('.') + 1)..]);
    }

#pragma warning disable CA1822
    public interface IThing
    {
        static abstract string Kind { get; }

        static virtual string Fallback => "fallback";

        static string Default => "default";

        string Name { get; }
    }

    public class Thing :
        IThing
    {
        public static string Kind => "kind";

        public string Name => "thing";
    }

    public ref struct Cursor
    {
        public int Position;
    }

    public class Wrapper<T>
    {
        public T Value { get; set; } = default!;
    }

    public class HiddenBase
    {
        public string Kept { get; set; } = "kept";

        public string ByMethod { get; set; } = "";

        public string ByEvent { get; set; } = "";

        public string ByType { get; set; } = "";

        public string ByWriteOnly { get; set; } = "";

        public string BySpan { get; set; } = "";
    }

    public class Hiding :
        HiddenBase
    {
        public new void ByMethod()
        {
        }

        public new event Action ByEvent
        {
            add { }
            remove { }
        }

        public new class ByType;

        public new string ByWriteOnly
        {
            set => Kept = value;
        }

        public new ReadOnlySpan<char> BySpan => Kept.AsSpan();
    }

    // Every member here is legal C#, and several are ones no accessor can be written for. That this
    // file compiles is the first assertion: the generator emits this model's accessors into it.
    [ParchmentBindable]
    public unsafe partial class Awkward
    {
        int count = 3;

        public string Name { get; set; } = "";

        public ref int Count => ref count;

        public ReadOnlySpan<char> Chars => Name.AsSpan();

        public Span<int> Numbers => default;

        public Cursor Cursor => default;

        public int* Pointer = null;

        public delegate*<int, int> Function = null;

        public int*[] Pointers { get; set; } = [];

        public List<int*[]> PointerLists { get; set; } = [];

        public IThing Thing { get; set; } = new Thing();

        public dynamic? Dyn { get; set; }

        public List<dynamic> Dyns { get; set; } = [];

        public Dictionary<string, dynamic> Map { get; set; } = [];

        public Wrapper<dynamic> Wrapped { get; set; } = new();

        public (int dynamic, dynamic B) Pair { get; set; }

        [ExcelsiorTable]
        public List<dynamic> Table { get; set; } = [];

        public Hiding Hiding { get; set; } = new();
    }
#pragma warning restore CA1822

    // A ref struct or a pointer cannot become an object, a static abstract interface member is
    // reachable through a type parameter only, and a name a derived type reuses for a method no
    // longer means the inherited property. None has an accessor, so none may register.
    [Test]
    [Arguments("Chars")]
    [Arguments("Numbers")]
    [Arguments("Cursor")]
    [Arguments("Pointer")]
    [Arguments("Function")]
    [Arguments("Pointers")]
    [Arguments("PointerLists")]
    [Arguments("Thing.Kind")]
    [Arguments("Thing.Fallback")]
    [Arguments("Hiding.ByMethod")]
    [Arguments("Hiding.ByEvent")]
    [Arguments("Hiding.ByType")]
    [Arguments("Hiding.ByWriteOnly")]
    public async Task UnbindableMember_FailsRegistration(string path)
    {
        using var template = DocxTemplateBuilder.Build($"{{{{ {path} }}}}");
        var exception = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterDocxTemplate<Awkward>(template));
        await Assert.That(exception.Message).Contains(path[(path.LastIndexOf('.') + 1)..]);

        var markdown = Assert.Throws<ParchmentRegistrationException>(
            () => new TemplateStore().RegisterMarkdownTemplate<Awkward>($"{{{{ {path} }}}}"));
        await Assert.That(markdown.Message).Contains(path[(path.LastIndexOf('.') + 1)..]);
    }

    [Test]
    public async Task MembersBesideUnbindableOnes_Render()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Name }}|{{ Count }}|{{ Dyn }}|{{ Wrapped.Value }}|{{ Hiding.Kept }}|{{ Thing.Name }}|{{ Thing.Default }}

            {% for item in Dyns %}

            item {{ item }}

            {% endfor %}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Awkward>(template);

        using var output = new MemoryStream();
        await store.Render(
            new Awkward
            {
                Name = "name",
                Dyn = 5,
                Dyns = ["one", 2],
                Wrapped =
                {
                    Value = "wrapped"
                }
            },
            output);

        output.Position = 0;
        using var document = WordprocessingDocument.Open(output, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .ToList();
        await Assert.That(paragraphs).IsEquivalentTo(
            [
                "name|3|5|wrapped|kept|thing|default",
                "item one",
                "item 2"
            ]);
    }

    [Test]
    public async Task InheritedMember_RegistersAndRenders()
    {
        using var template = DocxTemplateBuilder.Build("{{ Title }} — {{ Body }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Report>(template);

        using var output = new MemoryStream();
        await store.Render(
            new Report { Title = "Q1", Body = "summary" },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task ShadowedMember_DerivedWins()
    {
        // ShadowingReport hides DocumentBase.Title with `new`. Validation must see the derived
        // property (not error on collision) and rendering must pick up the derived value.
        using var template = DocxTemplateBuilder.Build("{{ Title }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShadowingReport>(template);

        using var output = new MemoryStream();
        await store.Render(
            new ShadowingReport { Title = "derived" },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task IfConditionWithUnknownMember_FailsRegistration()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {% if Bogus %}

            yes

            {% endif %}
            """);

        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<Doc>(template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("Bogus");
    }
}
