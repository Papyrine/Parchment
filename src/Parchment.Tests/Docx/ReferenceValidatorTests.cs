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
