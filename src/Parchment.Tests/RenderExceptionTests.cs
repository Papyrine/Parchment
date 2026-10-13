#pragma warning disable CA1822 // members that only throw still have to be instance members to bind
// ReSharper disable PartialTypeWithSinglePart
// A render reports every failure as a ParchmentRenderException naming the template and, where one
// is known, the part and the token or tag - never as whatever the model, a delegate or a library
// happened to throw.
public partial class RenderExceptionTests
{
    [ParchmentBindable]
    public partial class Model
    {
        public string Name { get; set; } = "Alice";
        public bool Flag { get; set; }
        public List<int> Numbers { get; set; } = [1, 2, 3];
        public long Big { get; set; } = 5_000_000_000;
        public TokenValue? Token { get; set; }

        public bool Thrower => throw new InvalidOperationException("boom-getter");
        public List<int> ThrowingList => throw new InvalidOperationException("boom-list");
        public bool Cancelled => throw new OperationCanceledException("stop");

        public IEnumerable<int> Lazy
        {
            get
            {
                yield return 1;
                throw new InvalidOperationException("boom-lazy");
            }
        }
    }

    public class Row
    {
        [Column(Order = 1)]
        public required string Name { get; init; }
    }

    public class ThrowingRow
    {
        [Column(Order = 1)]
        public string Name => throw new InvalidOperationException("boom-row");
    }

    [ParchmentBindable]
    public partial class ConfigureModel
    {
        [ExcelsiorTable(Configure = nameof(Configure))]
        public IReadOnlyList<Row> Rows { get; set; } = [new() { Name = "a" }];

        static void Configure(WordTableBuilder<Row> builder) =>
            throw new InvalidOperationException("boom-configure");
    }

    [ParchmentBindable]
    public partial class RowGetterModel
    {
        [ExcelsiorTable]
        public IReadOnlyList<ThrowingRow> Rows { get; set; } = [new()];
    }

    static async Task<ParchmentRenderException> DocxFailure<T>(string content, T model)
        where T : class
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<T>(template);

        using var stream = new MemoryStream();
        var exception = await Assert.That(() => store.Render(model, stream))
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.TemplateName).IsEqualTo(typeof(T).Name);
        await Assert.That(exception.PartUri).IsEqualTo("/word/document.xml");
        return exception;
    }

    static async Task<ParchmentRenderException> MarkdownFailure<T>(string markdown, T model)
        where T : class
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<T>(markdown, styleSource);

        using var stream = new MemoryStream();
        var exception = await Assert.That(() => store.Render(model, stream))
            .Throws<ParchmentRenderException>();
        await Assert.That(exception!.TemplateName).IsEqualTo(typeof(T).Name);
        return exception;
    }

    [Test]
    [Arguments("{% if Thrower %}\n\nx\n\n{% endif %}", "{% if Thrower %}", "boom-getter")]
    [Arguments("{% if Flag %}\n\nx\n\n{% elsif Thrower %}\n\ny\n\n{% endif %}", "{% elsif Thrower %}", "boom-getter")]
    [Arguments("{% for i in ThrowingList %}\n\n{{ i }}\n\n{% endfor %}", "{% for i in ThrowingList %}", "boom-list")]
    [Arguments("{% for i in Lazy %}\n\n{{ i }}\n\n{% endfor %}", "{% for i in Lazy %}", "boom-lazy")]
    [Arguments("{% for i in Numbers limit: Thrower %}\n\n{{ i }}\n\n{% endfor %}", "{% for i in Numbers limit: Thrower %}", "boom-getter")]
    [Arguments("{% for i in Numbers offset: Thrower %}\n\n{{ i }}\n\n{% endfor %}", "{% for i in Numbers offset: Thrower %}", "boom-getter")]
    public async Task Docx_BlockTagFailure_NamesTheTag(string content, string tag, string message)
    {
        var exception = await DocxFailure(content, new Model());

        await Assert.That(exception.TokenSource).IsEqualTo(tag);
        await Assert.That(exception.Message).Contains(message);
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    [Arguments("limit")]
    [Arguments("offset")]
    public async Task Docx_LoopModifierBeyondIntRange_SaysSo(string modifier)
    {
        var tag = $"{{% for i in Numbers {modifier}: Big %}}";
        var exception = await DocxFailure($"{tag}\n\n{{{{ i }}}}\n\n{{% endfor %}}", new Model());

        await Assert.That(exception.TokenSource).IsEqualTo(tag);
        await Assert.That(exception.Message).Contains($"'{modifier}' is 5000000000");
        await Assert.That(exception.InnerException).IsNull();
    }

    [Test]
    public async Task Docx_NestedFailure_IsWrappedOnce()
    {
        // The condition fails inside a loop inside a conditional; each of those catches.
        var exception = await DocxFailure(
            "{% if Name %}\n\n{% for i in Numbers %}\n\n{% if Thrower %}\n\nx\n\n{% endif %}\n\n{% endfor %}\n\n{% endif %}",
            new Model());

        await Assert.That(exception.TokenSource).IsEqualTo("{% if Thrower %}");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    [Arguments("{{ Token }}")]
    [Arguments("before {{ Token }} after")]
    [Arguments("{{ Name }} {{ Token }} after")]
    public async Task Docx_OpenXmlTokenDelegateThrows_NamesTheToken(string content)
    {
        var exception = await DocxFailure(
            content,
            new Model
            {
                Token = new OpenXmlToken(_ => throw new InvalidOperationException("boom-openxml"))
            });

        await Assert.That(exception.TokenSource).IsEqualTo("{{ Token }}");
        await Assert.That(exception.Message).Contains("boom-openxml");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    [Arguments("{{ Token }}")]
    [Arguments("{{ Name }} {{ Token }} after")]
    public async Task Docx_MutateTokenThrows_NamesTheToken(string content)
    {
        var exception = await DocxFailure(
            content,
            new Model
            {
                Token = new MutateToken((_, _) => throw new InvalidOperationException("boom-mutate"))
            });

        await Assert.That(exception.TokenSource).IsEqualTo("{{ Token }}");
        await Assert.That(exception.Message).Contains("boom-mutate");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    public async Task Docx_ExcelsiorConfigureThrows_SurfacesTheRealCause()
    {
        var exception = await DocxFailure("{{ Rows }}", new ConfigureModel());

        await Assert.That(exception.TokenSource).IsEqualTo("{{ Rows }}");
        await Assert.That(exception.Message).Contains("boom-configure");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    public async Task Docx_ExcelsiorRowGetterThrows_SurfacesTheRealCause()
    {
        var exception = await DocxFailure("{{ Rows }}", new RowGetterModel());

        await Assert.That(exception.TokenSource).IsEqualTo("{{ Rows }}");
        await Assert.That(exception.Message).Contains("boom-row");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    [Arguments("{{ Cancelled }}")]
    [Arguments("{% if Cancelled %}\n\nx\n\n{% endif %}")]
    [Arguments("{% for i in Numbers limit: Cancelled %}\n\n{{ i }}\n\n{% endfor %}")]
    public async Task Docx_Cancellation_IsNotWrapped(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Model>(template);

        using var stream = new MemoryStream();
        await Assert.That(() => store.Render(new Model(), stream))
            .Throws<OperationCanceledException>();
    }

    [Test]
    [Arguments("{{ Thrower }}", "boom-getter")]
    [Arguments("{% if Thrower %}x{% endif %}", "boom-getter")]
    [Arguments("{% for i in ThrowingList %}{{ i }}{% endfor %}", "boom-list")]
    [Arguments("{% for i in Lazy %}{{ i }}{% endfor %}", "boom-lazy")]
    [Arguments("{% for i in Numbers limit: Thrower %}{{ i }}{% endfor %}", "boom-getter")]
    public async Task Markdown_LiquidFailure_IsWrapped(string markdown, string message)
    {
        var exception = await MarkdownFailure(markdown, new Model());

        await Assert.That(exception.Message).Contains(message);
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    public async Task Markdown_LoopModifierBeyondIntRange_IsWrapped()
    {
        var exception = await MarkdownFailure("{% for i in Numbers limit: Big %}{{ i }}{% endfor %}", new Model());

        await Assert.That(exception.InnerException).IsTypeOf<OverflowException>();
    }

    [Test]
    public async Task Markdown_OpenXmlTokenDelegateThrows_IsWrapped()
    {
        var exception = await MarkdownFailure(
            "{{ Token }}",
            new Model
            {
                Token = new OpenXmlToken(_ => throw new InvalidOperationException("boom-openxml"))
            });

        await Assert.That(exception.Message).Contains("boom-openxml");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    public async Task Markdown_ExcelsiorConfigureThrows_SurfacesTheRealCause()
    {
        var exception = await MarkdownFailure("{{ Rows }}", new ConfigureModel());

        await Assert.That(exception.TokenSource).IsEqualTo("{{ Rows }}");
        await Assert.That(exception.Message).Contains("boom-configure");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    public async Task Markdown_ExcelsiorRowGetterThrows_SurfacesTheRealCause()
    {
        var exception = await MarkdownFailure("{{ Rows }}", new RowGetterModel());

        await Assert.That(exception.TokenSource).IsEqualTo("{{ Rows }}");
        await Assert.That(exception.Message).Contains("boom-row");
        await Assert.That(exception.InnerException).IsTypeOf<InvalidOperationException>();
    }

    [Test]
    public async Task Markdown_Cancellation_IsNotWrapped()
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Model>("{{ Cancelled }}", styleSource);

        using var stream = new MemoryStream();
        await Assert.That(() => store.Render(new Model(), stream))
            .Throws<OperationCanceledException>();
    }
}
