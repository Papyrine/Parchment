// ReSharper disable PartialTypeWithSinglePart
// RenderToFile only touches the destination once the render has succeeded: a failed or cancelled
// render leaves an existing document as it was, creates nothing at a new path, and leaves no
// temporary file behind.
public partial class RenderToFileTests
{
    [ParchmentBindable]
    public partial class Report
    {
        public string Title { get; init; } = "ok";
        public bool Fail { get; init; }

        public string Body
        {
            get
            {
                if (Fail)
                {
                    throw new InvalidOperationException("getter boom");
                }

                return "body";
            }
        }
    }

    [ParchmentBindable]
    public partial class Unregistered
    {
        public string Title { get; init; } = "ok";
    }

    static readonly byte[] original = [1, 2, 3, 4, 5];

    static TemplateStore Store()
    {
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Report>("# {{ Title }}\n\n{{ Body }}");
        return store;
    }

    static string NewDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"parchment-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        return directory;
    }

    static async Task<Exception?> Catch(Func<Task> action)
    {
        try
        {
            await action();
        }
        catch (Exception exception)
        {
            return exception;
        }

        return null;
    }

    static async Task AssertUntouched(string directory, string path)
    {
        await Assert.That(await File.ReadAllBytesAsync(path)).IsEquivalentTo(original);
        await Assert.That(Directory.GetFileSystemEntries(directory).Length).IsEqualTo(1);
    }

    [Test]
    public async Task FailedRenderKeepsTheExistingFile()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");
            await File.WriteAllBytesAsync(path, original);

            var exception = await Catch(() => Store().RenderToFile(new Report { Fail = true }, path));

            await Assert.That(exception).IsNotNull();
            await AssertUntouched(directory, path);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Test]
    public async Task FailedRenderCreatesNothingAtANewPath()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");

            var exception = await Catch(() => Store().RenderToFile(new Report { Fail = true }, path));

            await Assert.That(exception).IsNotNull();
            await Assert.That(Directory.GetFileSystemEntries(directory).Length).IsEqualTo(0);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Test]
    public async Task UnregisteredModelKeepsTheExistingFile()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");
            await File.WriteAllBytesAsync(path, original);

            var exception = await Catch(() => Store().RenderToFile(new Unregistered(), path));

            await Assert.That(exception).IsTypeOf<ParchmentRenderException>();
            await AssertUntouched(directory, path);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Test]
    public async Task NullModelKeepsTheExistingFile()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");
            await File.WriteAllBytesAsync(path, original);

            var exception = await Catch(() => Store().RenderToFile<Report>(null!, path));

            await Assert.That(exception).IsTypeOf<ParchmentRenderException>();
            await AssertUntouched(directory, path);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Test]
    public async Task CancelledRenderKeepsTheExistingFile()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");
            await File.WriteAllBytesAsync(path, original);
            using var source = new CancelSource();
            await source.CancelAsync();

            var exception = await Catch(() => Store().RenderToFile(new Report(), path, source.Token));

            await Assert.That(exception is OperationCanceledException).IsTrue();
            await AssertUntouched(directory, path);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Test]
    public async Task SuccessfulRenderReplacesTheExistingFile()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");
            await File.WriteAllBytesAsync(path, original);
            var store = Store();
            using var expected = new MemoryStream();
            await store.Render(new Report(), expected);

            await store.RenderToFile(new Report(), path);

            await Assert.That(await File.ReadAllBytesAsync(path)).IsEquivalentTo(expected.ToArray());
            await Assert.That(Directory.GetFileSystemEntries(directory).Length).IsEqualTo(1);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    [Test]
    public async Task RelativePathRendersIntoTheCurrentDirectory()
    {
        var name = $"parchment-{Guid.NewGuid():N}.docx";
        try
        {
            await Store().RenderToFile(new Report(), name);

            await Assert.That(new FileInfo(Path.Combine(Environment.CurrentDirectory, name)).Length).IsGreaterThan(0);
        }
        finally
        {
            File.Delete(name);
        }
    }

    [Test]
    public async Task MissingDirectoryThrowsDirectoryNotFound()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"parchment-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "report.docx");

        var exception = await Catch(() => Store().RenderToFile(new Report(), path));

        await Assert.That(exception).IsTypeOf<DirectoryNotFoundException>();
        await Assert.That(Directory.Exists(directory)).IsFalse();
    }

    [Test]
    public async Task LockedDestinationThrowsIoExceptionAndIsKept()
    {
        var directory = NewDirectory();
        try
        {
            var path = Path.Combine(directory, "report.docx");
            await File.WriteAllBytesAsync(path, original);

            Exception? exception;
            await using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                exception = await Catch(() => Store().RenderToFile(new Report(), path));
            }

            await Assert.That(exception).IsTypeOf<IOException>();
            await AssertUntouched(directory, path);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }
}
