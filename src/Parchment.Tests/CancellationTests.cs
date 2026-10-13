// ReSharper disable PartialTypeWithSinglePart

// The token a render is handed is checked inside the work, not only between the parts of the
// document: a loop stops at its next iteration, and nothing is read from the model for a token
// that was cancelled before the call.
public partial class CancellationTests
{
    const int rowCount = 2000;
    const int cancelAtRead = 10;

    const string loopTemplate =
        """
        {% for row in Rows %}

        {{ row.Value }}

        {% endfor %}
        """;

    [ParchmentBindable]
    public partial class Doc
    {
        public required IReadOnlyList<Row> Rows { get; init; }
    }

    // Counts what the render reads, and cancels the render's own token part-way through.
    public class Reads
    {
        public int Count { get; set; }
        public CancelSource? Source { get; init; }
        public bool Throw { get; init; }
    }

    public class Row(Reads reads)
    {
        public string Value
        {
            get
            {
                reads.Count++;
                if (reads.Count == cancelAtRead)
                {
                    reads.Source?.Cancel();
                    if (reads.Throw)
                    {
                        throw new OperationCanceledException("from the model");
                    }
                }

                return "x";
            }
        }
    }

    static Doc Model(Reads reads) =>
        new()
        {
            Rows = Enumerable.Range(0, rowCount).Select(_ => new Row(reads)).ToList()
        };

    static TemplateStore DocxStore(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return store;
    }

    static TemplateStore MarkdownStore(string markdown)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(markdown, styleSource);
        return store;
    }

    [Test]
    public async Task DocxLoopStopsAtTheNextIteration()
    {
        var store = DocxStore(loopTemplate);
        using var source = new CancelSource();
        var reads = new Reads
        {
            Source = source
        };

        using var output = new MemoryStream();
        await Assert.That(() => store.Render(Model(reads), output, source.Token))
            .ThrowsExactly<OperationCanceledException>();

        // The iteration that cancelled finishes its own token; no later one starts.
        await Assert.That(reads.Count).IsEqualTo(cancelAtRead);
        await Assert.That(output.Length).IsEqualTo(0);
    }

    [Test]
    public async Task DocxParagraphStopsAtTheNextToken()
    {
        // One paragraph and no loop: the read that cancels is the tenth, and the hundreds of
        // tokens still to run in the paragraph are never evaluated.
        var tokens = string.Concat(Enumerable.Range(0, 500).Select(_ => "{{ Rows[0].Value }}"));
        var store = DocxStore(tokens);
        using var source = new CancelSource();
        var reads = new Reads
        {
            Source = source
        };

        using var output = new MemoryStream();
        await Assert.That(() => store.Render(Model(reads), output, source.Token))
            .ThrowsExactly<OperationCanceledException>();
        await Assert.That(reads.Count).IsEqualTo(cancelAtRead);
    }

    [Test]
    public async Task DocxNestedLoopStopsAtTheNextIteration()
    {
        var store = DocxStore(
            """
            {% for outer in Rows limit:3 %}

            {% for row in Rows %}

            {{ row.Value }}

            {% endfor %}

            {% endfor %}
            """);
        using var source = new CancelSource();
        var reads = new Reads
        {
            Source = source
        };

        using var output = new MemoryStream();
        await Assert.That(() => store.Render(Model(reads), output, source.Token))
            .ThrowsExactly<OperationCanceledException>();
        await Assert.That(reads.Count).IsEqualTo(cancelAtRead);
    }

    [Test]
    public async Task CancellationThrownByTheModelIsNotWrapped()
    {
        var store = DocxStore(loopTemplate);
        var reads = new Reads
        {
            Throw = true
        };

        using var output = new MemoryStream();
        var exception = await Assert.That(() => store.Render(Model(reads), output))
            .ThrowsExactly<OperationCanceledException>();
        await Assert.That(exception!.Message).IsEqualTo("from the model");
    }

    [Test]
    public async Task MarkdownStopsAfterTheLiquidRender()
    {
        var store = MarkdownStore("{% for row in Rows limit:20 %}{{ row.Value }}{% endfor %}");
        using var source = new CancelSource();
        var reads = new Reads
        {
            Source = source
        };

        using var output = new MemoryStream();
        await Assert.That(() => store.Render(Model(reads), output, source.Token))
            .ThrowsExactly<OperationCanceledException>();
        await Assert.That(output.Length).IsEqualTo(0);
    }

    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task AlreadyCancelledTokenReadsNothing(bool markdown)
    {
        TemplateStore store;
        if (markdown)
        {
            store = MarkdownStore("{% for row in Rows %}{{ row.Value }}{% endfor %}");
        }
        else
        {
            store = DocxStore(loopTemplate);
        }

        using var source = new CancelSource();
        source.Cancel();
        var reads = new Reads();

        using var output = new MemoryStream();
        await Assert.That(() => store.Render(Model(reads), output, source.Token))
            .ThrowsExactly<OperationCanceledException>();
        await Assert.That(reads.Count).IsEqualTo(0);
        await Assert.That(output.Length).IsEqualTo(0);
    }

    [Test]
    public async Task CancellableTokenThatIsNeverCancelledRendersTheSameBytes()
    {
        // The Markdig walk only subscribes its check for a token that can be cancelled, so both
        // paths are run and have to agree.
        var store = MarkdownStore("# Title\n\n{% for row in Rows limit:5 %}- {{ row.Value }}\n{% endfor %}\n\n| a | b |\n|---|---|\n| 1 | 2 |");
        using var source = new CancelSource();

        using var plain = new MemoryStream();
        // ReSharper disable once MethodSupportsCancellation
        await store.Render(Model(new()), plain);
        using var cancellable = new MemoryStream();
        await store.Render(Model(new()), cancellable, source.Token);

        await Assert.That(cancellable.ToArray().SequenceEqual(plain.ToArray())).IsTrue();
    }
}
