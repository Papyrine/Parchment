// ReSharper disable PartialTypeWithSinglePart
// ReSharper disable InconsistentNaming
// ReSharper disable UnusedMember.Global
// ReSharper disable UnusedParameter.Local
// Lower-case type names are what is under test.
// ReSharper disable RedundantVerbatimPrefix
#pragma warning disable CS8981
#pragma warning disable CS8860

// The models below exist to be compiled: the generator runs over this project, so a name it
// writes without the escape the language needs, or two hint names it cannot tell apart, fails the
// build of the whole test project rather than a test. The tests then check that a template still
// uses the plain name.
namespace GeneratedNames.@namespace
{
    [ParchmentBindable]
    public partial class @event
    {
        public string @class { get; set; } = "";
        public string @string = "";
        public static string @default => "static";
        public @object @object { get; set; } = new();
        public List<string> @params { get; set; } = [];

        [EditableField]
        public string @new { get; set; } = "";

        [ExcelsiorTable(Configure = nameof(@static))]
        public List<@struct> @foreach { get; set; } = [];

        static void @static(WordTableBuilder<@struct> builder)
        {
        }
    }

    public class @object
    {
        [StringSyntax("html")]
        public string @namespace { get; set; } = "";

        [EditableField]
        public string @int { get; set; } = "";
    }

    public class @struct
    {
        public string @is { get; set; } = "";
    }

    public partial class @checked
    {
        [ParchmentBindable]
        public partial class @unchecked
        {
            public string @this { get; set; } = "";
        }
    }

    // Contextual keywords: legal bare as a member, not all of them legal bare as a type.
    [ParchmentBindable]
    public partial class @record
    {
        public string var { get; set; } = "";
        public string async { get; set; } = "";
        public string file { get; set; } = "";

        [EditableField]
        public string required { get; set; } = "";

        [EditableField]
        public string scoped { get; set; } = "";

        public @partial Nested { get; set; } = new();
    }

    [ParchmentBindable]
    public partial class @partial
    {
        [EditableField]
        public string record { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class async
    {
        public string Name { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class @file
    {
        public string Name { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class @required
    {
        public string Name { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class @scoped
    {
        public string Name { get; set; } = "";
    }

    // Hint names are compared without regard to case.
    [ParchmentBindable]
    public partial class Twin
    {
        public string Name { get; set; } = "";
    }

    [ParchmentBindable]
    public partial class twin
    {
        public string Name { get; set; } = "";
    }
}

// `A_B.C` and `A.B_C` flatten to one hint name, and to one accessor field name inside a model
// that reaches both.
namespace GeneratedNames.A_B
{
    [ParchmentBindable]
    public partial class C
    {
        public string Name { get; set; } = "";
    }
}

namespace GeneratedNames.A
{
    [ParchmentBindable]
    public partial class B_C
    {
        public string Name { get; set; } = "";
    }
}

namespace GeneratedNames
{
    using @namespace;

    [ParchmentBindable]
    public partial class Flattened
    {
        public A_B.C One { get; set; } = new();
        public A.B_C Two { get; set; } = new();
    }

    public class KeywordIdentifierTests
    {
        static @event BuildModel() =>
            new()
            {
                @class = "class value",
                @string = "field value",
                @object = new()
                {
                    @namespace = "<b>bold</b>",
                    @int = "nested editable"
                },
                @params = ["one", "two"],
                @new = "editable",
                @foreach =
                [
                    new()
                    {
                        @is = "cell"
                    }
                ]
            };

        [Test]
        public async Task KeywordMembersBindByTheirPlainName()
        {
            using var template = DocxTemplateBuilder.Build(
                """
                {{ class }}|{{ string }}|{{ default }}|{{ params.size }}|{{ object.int }}

                {{ object.namespace }}

                {{ new }}

                {{ foreach }}
                """);
            var store = new TemplateStore();
            store.RegisterDocxTemplate<@event>(template);

            var text = await RenderText(store, BuildModel());

            await Assert.That(text).Contains("class value|field value|static|2|nested editable");
            await Assert.That(text).Contains("bold");
            await Assert.That(text).DoesNotContain("<b>");
            await Assert.That(text).Contains("editable");
            await Assert.That(text).Contains("cell");
        }

        [Test]
        public async Task KeywordMembersRenderInMarkdown()
        {
            using var styleSource = DocxTemplateBuilder.Build();
            var store = new TemplateStore();
            store.RegisterMarkdownTemplate<@event>("{{ class }}|{{ object.int }}", styleSource);

            await Assert.That(await RenderText(store, BuildModel())).IsEqualTo("class value|nested editable");
        }

        [Test]
        public async Task KeywordEditableMembersRoundTrip()
        {
            using var template = DocxTemplateBuilder.Build(
                """
                {{ new }}

                {{ object.int }}
                """);
            var store = new TemplateStore();
            store.RegisterDocxTemplate<@event>(template);
            using var stream = new MemoryStream();
            await store.Render(BuildModel(), stream);
            stream.Position = 0;

            var result = ParchmentExtractor.Extract<@event>(stream);
            var target = new @event();
            result.ApplyTo(target);

            await Assert.That(result.Fields.Select(_ => _.Path)).IsEquivalentTo(["new", "object.int"]);
            await Assert.That(target.@new).IsEqualTo("editable");
            await Assert.That(target.@object.@int).IsEqualTo("nested editable");
        }

        [Test]
        public async Task ContextualKeywordNamesBind()
        {
            using var template = DocxTemplateBuilder.Build("{{ var }}|{{ async }}|{{ file }}|{{ Nested.record }}");
            var store = new TemplateStore();
            store.RegisterDocxTemplate<@record>(template);

            var text = await RenderText(
                store,
                new @record
                {
                    var = "a",
                    async = "b",
                    file = "c",
                    Nested = new()
                    {
                        record = "d"
                    }
                });

            await Assert.That(text).IsEqualTo("a|b|c|d");
        }

        // Each of these is registered by its own generated file; a hint-name clash used to leave
        // none of the project's models registered.
        [Test]
        public async Task ModelsWithClashingGeneratedNamesAreAllRegistered()
        {
            await Assert.That(await RenderName(new Twin { Name = "upper" })).IsEqualTo("upper");
            await Assert.That(await RenderName(new twin { Name = "lower" })).IsEqualTo("lower");
            await Assert.That(await RenderName(new A_B.C { Name = "first" })).IsEqualTo("first");
            await Assert.That(await RenderName(new A.B_C { Name = "second" })).IsEqualTo("second");
            await Assert.That(await RenderName(new async { Name = "async" })).IsEqualTo("async");
            await Assert.That(await RenderName(new @file { Name = "file" })).IsEqualTo("file");
            await Assert.That(await RenderName(new @required { Name = "required" })).IsEqualTo("required");
            await Assert.That(await RenderName(new @scoped { Name = "scoped" })).IsEqualTo("scoped");
        }

        [Test]
        public async Task NestedKeywordTypeIsRegistered()
        {
            using var template = DocxTemplateBuilder.Build("{{ this }}");
            var store = new TemplateStore();
            store.RegisterDocxTemplate<@checked.@unchecked>(template);

            var text = await RenderText(
                store,
                new @checked.@unchecked
                {
                    @this = "nested"
                });

            await Assert.That(text).IsEqualTo("nested");
        }

        [Test]
        public async Task MemberTypesWithClashingAccessorNamesBothBind()
        {
            using var template = DocxTemplateBuilder.Build("{{ One.Name }}|{{ Two.Name }}");
            var store = new TemplateStore();
            store.RegisterDocxTemplate<Flattened>(template);

            var text = await RenderText(
                store,
                new Flattened
                {
                    One = new()
                    {
                        Name = "one"
                    },
                    Two = new()
                    {
                        Name = "two"
                    }
                });

            await Assert.That(text).IsEqualTo("one|two");
        }

        static async Task<string> RenderName<T>(T model)
            where T : class
        {
            using var template = DocxTemplateBuilder.Build("{{ Name }}");
            var store = new TemplateStore();
            store.RegisterDocxTemplate<T>(template);
            return await RenderText(store, model);
        }

        static async Task<string> RenderText<T>(TemplateStore store, T model)
            where T : class
        {
            using var stream = new MemoryStream();
            await store.Render(model, stream);
            stream.Position = 0;

            using var doc = WordprocessingDocument.Open(stream, false);
            return doc.MainDocumentPart!.Document!.Body!.InnerText;
        }
    }
}
