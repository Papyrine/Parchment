using Microsoft.Extensions.Logging;

// The source generator reports these as PARCH019 at compile time, but only for a model carrying
// [ParchmentModel]. Hand registration never reaches the generator, so the runtime path has to say
// the same thing or the attribute is dropped in silence.
public class StaticRenderAttributeTests
{
    sealed class HtmlAttribute : Attribute;

    public class StaticHtmlModel
    {
        public required string Body { get; init; }

        [Html]
        public static string Banner { get; set; } = "<b>banner</b>";
    }

    public class InstanceHtmlModel
    {
        public required string Body { get; init; }

        [Html]
        public required string Banner { get; init; }
    }

    public class PlainStaticModel
    {
        public required string Body { get; init; }

        public static string Banner { get; set; } = "plain";
    }

    public class StaticExcelsiorModel
    {
        public required string Body { get; init; }

        [ExcelsiorTable]
        public static IReadOnlyList<Row> Rows { get; set; } = [];
    }

    public class StaticStringSyntaxModel
    {
        public required string Body { get; init; }

        [StringSyntax("html")]
        public static string Banner { get; set; } = "<b>banner</b>";
    }

    public class NestedOwnerModel
    {
        public required string Body { get; init; }
        public required Nested Child { get; init; }
    }

    public class Nested
    {
        public required string Value { get; init; }

        [Html]
        public static string Banner { get; set; } = "<b>banner</b>";
    }

    public class Row
    {
        public required string Name { get; init; }
    }

    sealed class RecordingLogger : ILogger<TemplateStore>
    {
        public List<string> Warnings { get; } = [];

        public IDisposable? BeginScope<TState>(TState state)
            where TState : notnull =>
            null;

        public bool IsEnabled(LogLevel logLevel) => true;

        public void Log<TState>(
            LogLevel logLevel,
            EventId eventId,
            TState state,
            Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            if (logLevel == LogLevel.Warning)
            {
                Warnings.Add(formatter(state, exception));
            }
        }
    }

    static RecordingLogger RegisterDocx<TModel>()
    {
        var logger = new RecordingLogger();
        using var template = DocxTemplateBuilder.Build("{{ Body }}");
        new TemplateStore(logger).RegisterDocxTemplate<TModel>("t", template);
        return logger;
    }

    [Test]
    public async Task StaticHtmlMemberWarns()
    {
        var logger = RegisterDocx<StaticHtmlModel>();

        await Assert.That(logger.Warnings.Count).IsEqualTo(1);
        await Assert.That(logger.Warnings[0]).Contains("Html");
        await Assert.That(logger.Warnings[0]).Contains("StaticHtmlModel.Banner");
    }

    [Test]
    public async Task StaticExcelsiorTableMemberWarns()
    {
        var logger = RegisterDocx<StaticExcelsiorModel>();

        await Assert.That(logger.Warnings.Count).IsEqualTo(1);
        await Assert.That(logger.Warnings[0]).Contains("ExcelsiorTable");
    }

    // [StringSyntax] is a written directive like the others, and the generator's shape reads it the
    // same way, so a static one has to warn here too or the two paths disagree.
    [Test]
    public async Task StaticStringSyntaxMemberWarns()
    {
        var logger = RegisterDocx<StaticStringSyntaxModel>();

        await Assert.That(logger.Warnings.Count).IsEqualTo(1);
        await Assert.That(logger.Warnings[0]).Contains("StringSyntax");
    }

    // The generator walks every type in the model shape, not only the root.
    [Test]
    public async Task StaticMemberOnANestedTypeWarns()
    {
        var logger = RegisterDocx<NestedOwnerModel>();

        await Assert.That(logger.Warnings.Count).IsEqualTo(1);
        await Assert.That(logger.Warnings[0]).Contains("Nested.Banner");
    }

    [Test]
    public async Task InstanceMemberDoesNotWarn()
    {
        var logger = RegisterDocx<InstanceHtmlModel>();

        await Assert.That(logger.Warnings).IsEmpty();
    }

    // Only an attribute is evidence of a mistake. A plain static member is left alone, and binds
    // and renders normally.
    [Test]
    public async Task PlainStaticMemberDoesNotWarn()
    {
        var logger = RegisterDocx<PlainStaticModel>();

        await Assert.That(logger.Warnings).IsEmpty();
    }

    [Test]
    public async Task MarkdownRegistrationWarnsToo()
    {
        var logger = new RecordingLogger();
        using var styleSource = DocxTemplateBuilder.Build();
        new TemplateStore(logger).RegisterMarkdownTemplate<StaticHtmlModel>("t", "{{ Body }}", styleSource);

        await Assert.That(logger.Warnings.Count).IsEqualTo(1);
        await Assert.That(logger.Warnings[0]).Contains("StaticHtmlModel.Banner");
    }
}
