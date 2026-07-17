public class BindingScenarioTests
{
    public class DotxModel
    {
        public required string Title { get; init; }
    }

    // RegisterDocxTemplate clones the supplied package and the clone becomes the output, so a .dotx
    // template produced a template-typed package that Word opened as a new unsaved document.
    [Test]
    public async Task DotxTemplateProducesDocumentTypedOutput()
    {
        using var dotx = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(dotx, WordprocessingDocumentType.Template))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new(new Body(new Paragraph(new Run(new Text("{{ Title }}")))));
        }

        dotx.Position = 0;

        var store = new TemplateStore();
        store.RegisterDocxTemplate<DotxModel>("dotx", dotx);

        using var output = new MemoryStream();
        await store.Render(
            "dotx",
            new DotxModel
            {
                Title = "x"
            },
            output);
        output.Position = 0;

        using var result = WordprocessingDocument.Open(output, false);
        await Assert.That(result.DocumentType).IsEqualTo(WordprocessingDocumentType.Document);
    }

    // -- #1 public fields ----------------------------------------------------

    public class FieldModel
    {
        public string Title = "";
        public NestedField Nested = new();
    }

    public class NestedField
    {
        public string Value = "";
    }

    [Test]
    public async Task PublicField_RendersLikeProperty()
    {
        using var template = DocxTemplateBuilder.Build("Title: {{ Title }} — Nested: {{ Nested.Value }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<FieldModel>("field", template);

        using var output = new MemoryStream();
        await store.Render(
            "field",
            new FieldModel { Title = "T", Nested = new() { Value = "N" } },
            output);
        await Verify(output, "docx");
    }

    // -- #2 polymorphism / abstract base ------------------------------------

    public abstract class DocumentBase
    {
        public required string Title { get; init; }
    }

    public sealed class ConcreteReport : DocumentBase
    {
        public required string Body { get; init; }
    }

    [Test]
    public async Task AbstractBase_BindsAgainstDeclaredType_AcceptsSubclassInstance()
    {
        // Register against the abstract base. `Title` is on the base, so it must validate and
        // render. The instance passed in is a concrete subclass — `Render`'s IsInstanceOfType
        // accepts it. Subclass-only members (Body) are deliberately NOT referenced here; that
        // case is covered by the negative test below.
        using var template = DocxTemplateBuilder.Build("Doc: {{ Title }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<DocumentBase>("polymorphic", template);

        using var output = new MemoryStream();
        await store.Render(
            "polymorphic",
            new ConcreteReport { Title = "Q1", Body = "ignored" },
            output);
        await Verify(output, "docx");
    }

    [Test]
    public async Task AbstractBase_SubclassOnlyMember_FailsValidation()
    {
        // The token references `Body`, declared only on ConcreteReport. Registering against the
        // abstract base must reject the template — the contract is "bind against the declared
        // type", not "bind against any subclass you might pass at render time".
        using var template = DocxTemplateBuilder.Build("Doc: {{ Title }} — {{ Body }}");
        var store = new TemplateStore();
        var exception = await Assert.That(
                () => store.RegisterDocxTemplate<DocumentBase>("polymorphic-bad", template))
            .Throws<ParchmentRegistrationException>();
        await Assert.That(exception!.Message).Contains("Body");
    }

    // -- static members ------------------------------------------------------

    public class StaticHostModel
    {
        public static string Logo { get; } = "Acme Inc.";
        public static string Footer = "© 2026";

        public required string Name { get; init; }
    }

    [Test]
    public async Task StaticMembers_BindAtRoot()
    {
        using var template = DocxTemplateBuilder.Build("{{ Logo }} — {{ Name }} — {{ Footer }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<StaticHostModel>("statics", template);

        using var output = new MemoryStream();
        await store.Render(
            "statics",
            new StaticHostModel
            {
                Name = "Q1"
            },
            output);
        await Verify(output, "docx");
    }

    // -- #3 record struct ---------------------------------------------------

    public record struct Point(int X, int Y);

    public class ShapeModel
    {
        public required string Name { get; init; }
        public required Point Origin { get; init; }
    }

    [Test]
    public async Task RecordStruct_NestedMember_BindsAndRenders()
    {
        using var template = DocxTemplateBuilder.Build("{{ Name }} @ ({{ Origin.X }}, {{ Origin.Y }})");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<ShapeModel>("record-struct", template);

        using var output = new MemoryStream();
        await store.Render(
            "record-struct",
            new ShapeModel { Name = "origin", Origin = new(3, 7) },
            output);
        await Verify(output, "docx");
    }

    // -- #4 null model ------------------------------------------------------

    public class TitledModel
    {
        public required string Title { get; init; }
    }

    [Test]
    public async Task NullModel_ThrowsClean_NotNullReferenceException()
    {
        // Render(null) must surface a clean ParchmentRenderException naming the template, not a
        // raw NullReferenceException from string-interpolating model.GetType().Name into the
        // type-mismatch message.
        using var template = DocxTemplateBuilder.Build("{{ Title }}");
        var store = new TemplateStore();
        store.RegisterDocxTemplate<TitledModel>("null-model", template);

        using var output = new MemoryStream();
        var exception = await Assert.That(
                () => store.Render("null-model", null!, output))
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.Message).Contains("null");
    }

    // -- #5 IDictionary<K, V> member access through KeyValuePair -----------

    public class DictionaryHost
    {
        public required IReadOnlyDictionary<string, ProductInfo> Products { get; init; }
    }

    public class ProductInfo
    {
        public required string Sku { get; init; }
        public required string Description { get; init; }
    }

    [Test]
    public async Task DictionaryProperty_IteratedAsKvp_ResolvesValueMembers()
    {
        // Iterating an IDictionary<string, ProductInfo> in liquid yields key/value pairs where
        // `kv[0]` is the key and `kv[1]` is the value (Shopify-style; Fluid mirrors it). For
        // `kv[1].Sku` to render, ProductInfo's members must be registered with Fluid's strategy.
        // SharedFluid.Unwrap used to return the FIRST generic argument of any generic IEnumerable
        // (so `Dictionary<string, ProductInfo>` resolved to `string`), dropping ProductInfo from
        // the type-graph walk and silently rendering Sku empty.
        using var template = DocxTemplateBuilder.Build(
            """
            {% for kv in Products %}

            {{ kv[1].Sku }}={{ kv[1].Description }}

            {% endfor %}
            """);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<DictionaryHost>("dict-host", template);

        using var output = new MemoryStream();
        await store.Render(
            "dict-host",
            new DictionaryHost
            {
                Products = new Dictionary<string, ProductInfo>
                {
                    ["a"] = new() { Sku = "S1", Description = "Widget" },
                    ["b"] = new() { Sku = "S2", Description = "Gadget" }
                }
            },
            output);
        output.Position = 0;

        using var doc = WordprocessingDocument.Open(output, false);
        var paragraphs = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .Select(_ => _.InnerText)
            .Where(_ => _.Length > 0)
            .ToList();

        await Assert.That(paragraphs).Contains("S1=Widget");
        await Assert.That(paragraphs).Contains("S2=Gadget");
    }
}
