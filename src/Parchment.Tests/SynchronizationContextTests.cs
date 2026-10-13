public partial class SynchronizationContextTests
{
    [ParchmentBindable]
    public partial class Letter
    {
        public string Title { get; set; } = "Title";
    }

    // A UI thread or a classic ASP.NET request: one thread, and a continuation posted to the
    // context runs only when that thread comes back to pump it.
    sealed class SingleThreadContext :
        SynchronizationContext
    {
        public readonly ConcurrentQueue<SendOrPostCallback> Posted = new();

        public override void Post(SendOrPostCallback callback, object? state) =>
            Posted.Enqueue(callback);
    }

    // A network or asynchronous file stream: the write completes later, on another thread, so the
    // awaiting render has to choose where to resume. Not a MemoryStream, which CopyToAsync writes
    // to synchronously.
    sealed class AsyncWriteStream :
        Stream
    {
        long written;

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => written;

        public override long Position
        {
            get => written;
            set => throw new NotSupportedException();
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();

        public override void SetLength(long value) =>
            throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) =>
            written += count;

        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, Cancel cancel = default)
        {
            await Task.Delay(5, cancel).ConfigureAwait(false);
            written += buffer.Length;
        }

        public override async Task WriteAsync(byte[] buffer, int offset, int count, Cancel cancel)
        {
            await Task.Delay(5, cancel).ConfigureAwait(false);
            written += count;
        }
    }

    // Blocking on a render from a single-threaded context deadlocks when any await inside the
    // library resumes on the captured context: the continuation is posted to the one thread that
    // is blocked waiting for it. The wait is bounded so a regression fails instead of hanging.
    [Test]
    [Arguments("docx", "Render")]
    [Arguments("docx", "RenderToFile")]
    [Arguments("markdown", "Render")]
    [Arguments("markdown", "RenderToFile")]
    public async Task BlockingRenderDoesNotDeadlock(string flow, string api)
    {
        var store = new TemplateStore();
        if (flow == "docx")
        {
            using var template = DocxTemplateBuilder.Build("{{ Title }}");
            store.RegisterDocxTemplate<Letter>(template);
        }
        else
        {
            store.RegisterMarkdownTemplate<Letter>("# {{ Title }}");
        }

        var path = Path.Combine(Path.GetTempPath(), $"parchment-context-{Guid.NewGuid():N}.docx");
        var completed = false;
        var posted = 0;
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            var context = new SingleThreadContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                using var output = new AsyncWriteStream();
                var task = Start(store, api, output, path);
                completed = task.Wait(TimeSpan.FromSeconds(30));
                posted = context.Posted.Count;
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });
        thread.Start();
        thread.Join();

        try
        {
            await Assert.That(failure).IsNull();
            await Assert.That(completed).IsTrue();
            await Assert.That(posted).IsEqualTo(0);
        }
        finally
        {
            if (completed)
            {
                File.Delete(path);
            }
        }
    }

    static Task Start(TemplateStore store, string api, Stream output, string path)
    {
        if (api == "Render")
        {
            return store.Render(new Letter(), output);
        }

        return store.RenderToFile(new Letter(), path);
    }
}
