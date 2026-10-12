// ReSharper disable PartialTypeWithSinglePart
public partial class StringListTests
{
    static string ScenarioPath(string scenarioName) =>
        Path.GetFullPath(Path.Combine(
            ProjectFiles.ProjectDirectory,
            "Scenarios",
            scenarioName));

    #region StringListModel
    [ParchmentBindable]
    public partial class Person
    {
        public required string Name;
        public required IEnumerable<string> Tags;
    }
    #endregion

    [ParchmentBindable]
    public partial class PersonWithArray
    {
        public required string Name { get; init; }
        public required string[] Tags { get; init; }
    }

    [ParchmentBindable]
    public partial class PersonWithList
    {
        public required string Name { get; init; }
        public required List<string> Tags { get; init; }
    }

    [ParchmentBindable]
    public partial class PersonWithReadOnlyList
    {
        public required string Name { get; init; }
        public required IReadOnlyList<string> Tags { get; init; }
    }

    [ParchmentBindable]
    public partial class PersonWithNullableTags
    {
        public required string Name { get; init; }
        public IEnumerable<string>? Tags { get; init; }
    }

    [ParchmentBindable]
    public partial class Customer
    {
        public required string Name { get; init; }
        public required Profile Profile { get; init; }
    }

    public class Profile
    {
        public required IReadOnlyList<string> Skills { get; init; }
    }

    [ParchmentBindable]
    public partial class TwoListsDoc
    {
        public required IReadOnlyList<string> First { get; init; }
        public required IReadOnlyList<string> Second { get; init; }
    }

    [ParchmentBindable]
    public partial class CustomersDoc
    {
        public required IReadOnlyList<CustomerWithTags> Customers { get; init; }
    }

    public class CustomerWithTags
    {
        public required string Name { get; init; }
        public required IReadOnlyList<string> Tags { get; init; }
    }

    [ParchmentBindable]
    public partial class PersonWithTagsField
    {
        public required string Name { get; init; }
        public IReadOnlyList<string> Tags = [];
    }

    [Test]
    public async Task FieldStringListRendersAsBulletList()
    {
        using var template = DocxTemplateBuilder.Build("{{ Tags }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<PersonWithTagsField>(template);

        var model = new PersonWithTagsField
        {
            Name = "Ada",
            Tags = ["a", "b"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test, Explicit]
    public async Task GenerateScenarioInputDocx()
    {
        // One-shot generator for the scenario input.docx. Run explicitly when authoring the
        // scenario; the resulting file is committed alongside the README assets.
        using var template = DocxTemplateBuilder.Build(
            """
            Person: {{ Name }}

            Tags:

            {{ Tags }}

            End.
            """);

        var path = Path.Combine(ScenarioPath("string-list"), "input.docx");
        await using var file = File.Create(path);
        template.Position = 0;
        await template.CopyToAsync(file);
    }

    [Test]
    public async Task Render()
    {
        #region StringListUsage

        var templatePath = Path.Combine(ScenarioPath("string-list"), "input.docx");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Person>(templatePath);

        var model = new Person
        {
            Name = "Ada Lovelace",
            Tags = ["Author", "Mathematician", "Engineer"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);

        #endregion

        var settings = new VerifySettings();
        settings.UseDirectory(ScenarioPath("string-list"));
        settings.UseFileName("output");

        stream.Position = 0;
        await Verify(stream, "docx", settings);
    }

    [Test]
    public async Task TopLevelEnumerableStringRendersAsBulletList()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Person: {{ Name }}

            {{ Tags }}

            End.
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Person>(template);

        var model = new Person
        {
            Name = "Ada",
            Tags = ["Author", "Mathematician", "Engineer"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task StringArrayRendersAsBulletList()
    {
        using var template = DocxTemplateBuilder.Build("{{ Tags }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<PersonWithArray>(template);

        var model = new PersonWithArray
        {
            Name = "Ada",
            Tags = ["a", "b", "c"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task ListOfStringRendersAsBulletList()
    {
        using var template = DocxTemplateBuilder.Build("{{ Tags }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<PersonWithList>(template);

        var model = new PersonWithList
        {
            Name = "Ada",
            Tags = ["a", "b"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task ReadOnlyListOfStringRendersAsBulletList()
    {
        using var template = DocxTemplateBuilder.Build("{{ Tags }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<PersonWithReadOnlyList>(template);

        var model = new PersonWithReadOnlyList
        {
            Name = "Ada",
            Tags = ["a", "b"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task NullCollectionRendersAsEmpty()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            Header

            {{ Tags }}

            Footer
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<PersonWithNullableTags>(template);

        var model = new PersonWithNullableTags
        {
            Name = "Ada",
            Tags = null
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task NestedDottedPathRendersAsBulletList()
    {
        using var template = DocxTemplateBuilder.Build(
            """
            {{ Name }}

            {{ Profile.Skills }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Customer>(template);

        var model = new Customer
        {
            Name = "Bob",
            Profile = new()
            {
                Skills = ["C#", "OpenXML", "Liquid"]
            }
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task LoopScopeStringListFallsThroughToFluid()
    {
        // {{ c.Tags }} inside a {% for c in Customers %} loop is keyed on a loop variable,
        // not the root model — the auto-bullet-list path skips loop scope, matching the
        // existing Excelsior/Format limitation. Inside a loop the user opts in explicitly
        // with the `bullet_list` filter. Render must succeed.
        using var template = DocxTemplateBuilder.Build(
            """
            Customers

            {% for c in Customers %}

            Name: {{ c.Name }}

            {{ c.Tags | bullet_list }}

            {% endfor %}

            End
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<CustomersDoc>(template);

        var model = new CustomersDoc
        {
            Customers =
            [
                new()
                {
                    Name = "Ada",
                    Tags = ["a", "b"]
                }
            ]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task MixedInlineStringListFallsThroughToFluid()
    {
        // The auto-bullet path requires solo-in-paragraph (otherwise the structural swap would
        // drop the prefix). With a sibling text/token, the auto path skips silently and Fluid
        // handles the substitution instead — same behavior as before this feature existed.
        using var template = DocxTemplateBuilder.Build("Prefix {{ Tags }}");

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Person>(template);

        var model = new Person
        {
            Name = "Ada",
            Tags = ["a", "b"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [Test]
    public async Task TwoStringListsShareOneBulletAbstract()
    {
        // Two IEnumerable<string> properties on the same model must produce a single shared
        // bullet abstract — not one per token. Each list created its own WordNumberingState
        // before, which (a) wasted definitions and (b) gave non-deterministic abstract IDs
        // because OOXML SDK reorders abstracts based on document position when both are
        // inserted via InsertAt(0).
        using var template = DocxTemplateBuilder.Build(
            """
            First:

            {{ First }}

            Second:

            {{ Second }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<TwoListsDoc>(template);

        var model = new TwoListsDoc
        {
            First = ["a1", "a2"],
            Second = ["b1", "b2"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var bulletAbstracts = numbering
            .Elements<AbstractNum>()
            .Where(a =>
                a.Elements<Level>().FirstOrDefault(_ => _.LevelIndex?.Value == 0) is { } lvl &&
                lvl.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
            .ToList();
        var instances = numbering.Elements<NumberingInstance>().ToList();
        var instancesPointingAtBulletAbstract = instances
            .Where(i =>
                i.AbstractNumId?.Val?.Value is { } refId &&
                bulletAbstracts.Any(a => a.AbstractNumberId?.Value == refId))
            .ToList();

        await Assert.That(bulletAbstracts.Count).IsEqualTo(1);
        await Assert.That(instancesPointingAtBulletAbstract.Count).IsEqualTo(2);
    }

    [Test]
    public async Task EmptyCollectionRendersWithoutListParagraphs()
    {
        // An empty IEnumerable<string> still triggers the auto-bullet path (gates pass: solo,
        // plain identifier, IEnumerable<string>) but produces zero list paragraphs. The
        // surrounding paragraphs ("Header"/"Footer") must remain.
        using var template = DocxTemplateBuilder.Build(
            """
            Header

            {{ Tags }}

            Footer
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Person>(template);

        var model = new Person
        {
            Name = "Ada",
            Tags = []
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();

        var listParagraphs = paragraphs.Where(_ =>
                _.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph")
            .ToList();
        await Assert.That(listParagraphs).IsEmpty();

        var visibleTexts = paragraphs
            .Select(_ => string.Concat(_.Descendants<Text>().Select(t => t.Text)))
            .Where(_ => _.Length > 0)
            .ToList();
        await Assert.That(visibleTexts).IsEquivalentTo(["Header", "Footer"]);
    }

    [Test]
    public async Task FilterOnStringListFallsThroughToFluid()
    {
        // Applying a filter is an explicit opt-out of the auto-bullet path. The existing
        // bullet_list / numbered_list filters still work on IEnumerable<string> properties.
        using var template = DocxTemplateBuilder.Build(
            """
            Tags:

            {{ Tags | numbered_list }}
            """);

        var store = new TemplateStore();
        store.RegisterDocxTemplate<Person>(template);

        var model = new Person
        {
            Name = "Ada",
            Tags = ["first", "second", "third"]
        };

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        await Verify(stream, "docx");
    }

    [ParchmentBindable]
    public partial class ListShapes
    {
        public string[] Array { get; init; } = [];
        public List<string> List { get; init; } = [];
        public IEnumerable<string> Enumerable { get; init; } = [];
        public IReadOnlyList<string> ReadOnlyList { get; init; } = [];
        public HashSet<string> Set { get; init; } = [];
        public string[][] Jagged { get; init; } = [];
        public string[,] Grid { get; init; } = new string[0, 0];
    }

    static async Task<List<(string Text, bool IsListItem)>> RenderShape(string template, ListShapes model)
    {
        using var input = DocxTemplateBuilder.Build(template);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<ListShapes>(input);

        using var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;

        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Select(_ => (_.InnerText, _.ParagraphProperties?.NumberingProperties != null))
            .ToList();
    }

    [Test]
    [Arguments("Array")]
    [Arguments("List")]
    [Arguments("Enumerable")]
    [Arguments("ReadOnlyList")]
    [Arguments("Set")]
    public async Task EveryStringSequenceShapeRendersAsBulletList(string member)
    {
        var model = new ListShapes
        {
            Array = ["a", "b"],
            List = ["a", "b"],
            Enumerable = ["a", "b"],
            ReadOnlyList = ["a", "b"],
            Set = ["a", "b"]
        };

        var paragraphs = await RenderShape($"{{{{ {member} }}}}", model);

        await Assert.That(paragraphs).IsEquivalentTo([("a", true), ("b", true)]);
    }

    // A rank-2 array is not an IEnumerable<string>. It used to be classified as a string list
    // anyway, so the solo-token path replaced its paragraph with a list of no items and the member
    // vanished from the document without a word. It is an ordinary member now: the token reaches
    // Fluid alone in its paragraph exactly as it does with text beside it.
    [Test]
    [Arguments("{{ Grid }}")]
    [Arguments("Grid: {{ Grid }}")]
    public async Task MultiDimensionalArrayIsNotAStringList(string template)
    {
        var model = new ListShapes
        {
            Grid = new[,] { { "a", "b" }, { "c", "d" } }
        };

        var exception = await Assert.That(async () => { await RenderShape(template, model); })
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.Message).Contains("{{ Grid }}");
    }

    // A jagged array is a sequence of string[], never of string, so it was never a string list and
    // is not one now: no list paragraphs, and the rows print the way Fluid prints nested arrays.
    [Test]
    public async Task JaggedArrayIsNotAStringList()
    {
        var model = new ListShapes
        {
            Jagged = [["a", "b"], ["c"]]
        };

        var paragraphs = await RenderShape("{{ Jagged }}", model);

        await Assert.That(paragraphs).IsEquivalentTo([("abc", false)]);
    }
}
