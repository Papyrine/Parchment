// ReSharper disable PartialTypeWithSinglePart
public partial class DateTimeTokenTests
{
    [ParchmentBindable]
    public partial class Doc
    {
        public DateTime Unspecified { get; init; } = new(2026, 3, 1, 13, 5, 9);
        public DateTime Utc { get; init; } = new(2026, 3, 1, 13, 5, 9, DateTimeKind.Utc);
        public DateTime Local { get; init; } = new(2026, 3, 1, 13, 5, 9, DateTimeKind.Local);
        public DateTime? Nullable { get; init; } = new(2026, 3, 1, 13, 5, 9);
        public DateTimeOffset Offset { get; init; } = new(2026, 3, 1, 13, 5, 9, TimeSpan.FromHours(5));
        public DateOnly Day { get; init; } = new(2026, 3, 1);
        public string Text { get; init; } = "2026-03-01 13:05:09";
        public long Unix { get; init; } = 1772370309;

        // Ordered by wall clock. Read as local times, the two Unspecified entries land on either
        // side of the Utc one depending on the machine's zone.
        public IReadOnlyList<Entry> Entries { get; init; } =
        [
            new()
            {
                Name = "c",
                When = new(2026, 3, 1, 13, 0, 0)
            },
            new()
            {
                Name = "b",
                When = new(2026, 3, 1, 12, 30, 0, DateTimeKind.Utc)
            },
            new()
            {
                Name = "a",
                When = new(2026, 3, 1, 12, 0, 0)
            }
        ];
    }

    public class Entry
    {
        public required string Name { get; init; }
        public required DateTime When { get; init; }
    }

    const string stamp = "'%Y-%m-%d %H:%M:%S %z'";

    // Token to the text it prints on every machine. Each renders in both flows.
    static readonly (string Token, string Expected)[] zoneFree =
    [
        ("{{ Unspecified }}", "2026-03-01 13:05:09Z"),
        ("{{ Nullable }}", "2026-03-01 13:05:09Z"),
        ("{{ Utc }}", "2026-03-01 13:05:09Z"),
        // An offset-bearing value is an instant, and the bare token prints that instant in UTC.
        ("{{ Offset }}", "2026-03-01 08:05:09Z"),
        ("{{ Day }}", "03/01/2026"),
        ("{{ Unspecified | date }}", "2026-03-01 13:05:09Z"),
        ($"{{{{ Unspecified | date: {stamp} }}}}", "2026-03-01 13:05:09 +0000"),
        ($"{{{{ Utc | date: {stamp} }}}}", "2026-03-01 13:05:09 +0000"),
        // With a format the value's own wall clock and offset are printed.
        ($"{{{{ Offset | date: {stamp} }}}}", "2026-03-01 13:05:09 +0500"),
        ("{{ Unspecified | date: '%s' }}", "1772370309"),
        ("{{ Day | date: '%Y-%m-%d' }}", "2026-03-01"),
        // A string or a unix timestamp handed to `date` is read as UTC, not as machine time.
        ($"{{{{ Text | date: {stamp} }}}}", "2026-03-01 13:05:09 +0000"),
        ($"{{{{ Unix | date: {stamp} }}}}", "2026-03-01 13:05:09 +0000"),
        ("{{ Entries | sort: 'When' | map: 'Name' | join: ',' }}", "a,b,c")
    ];

    static readonly string[] zones =
    [
        "UTC",
        "Tokyo Standard Time",
        "Pacific Standard Time",
        "AUS Eastern Standard Time"
    ];

    [Test]
    public async Task DateTokensPrintTheSameTextInBothFlows()
    {
        foreach (var (token, expected) in zoneFree)
        {
            await Assert.That(await RenderDocx(token)).IsEqualTo(expected).Because(token);
            await Assert.That(await RenderMarkdown(token)).IsEqualTo(expected).Because(token);
        }
    }

    [Test]
    public async Task UnspecifiedEqualsUtcOfTheSameWallClock()
    {
        var docx = await RenderDocx(
            """
            {% if Unspecified == Utc %}

            same

            {% else %}

            differ

            {% endif %}
            """);
        var markdown = await RenderMarkdown("{% if Unspecified == Utc %}same{% else %}differ{% endif %}");

        await Assert.That(docx).IsEqualTo("same");
        await Assert.That(markdown).IsEqualTo("same");
    }

    // The machine's zone is process-wide state, so the tests that swap it run on their own. On a
    // machine set to UTC the tests above cannot tell a zone-dependent render from a correct one;
    // these can.
    [Test]
    [NotInParallel]
    public async Task DateTokensIgnoreTheMachineTimeZone()
    {
        foreach (var zone in zones)
        {
            using (LocalTimeZone.Emulate(zone))
            {
                foreach (var (token, expected) in zoneFree)
                {
                    await Assert.That(await RenderDocx(token)).IsEqualTo(expected).Because($"{token} in {zone}");
                    await Assert.That(await RenderMarkdown(token)).IsEqualTo(expected).Because($"{token} in {zone}");
                }
            }
        }
    }

    [Test]
    [NotInParallel]
    public async Task RenderIsByteIdenticalAcrossMachineTimeZones()
    {
        const string content = "{{ Unspecified }} {{ Text | date: '%H:%M %z' }} {{ Unix | date: '%H:%M %z' }}";
        using var template = DocxTemplateBuilder.Build(content);
        using var styleSource = DocxTemplateBuilder.Build();
        var docx = new HashSet<string>();
        var markdown = new HashSet<string>();
        foreach (var zone in zones)
        {
            using (LocalTimeZone.Emulate(zone))
            {
                var docxStore = new TemplateStore();
                template.Position = 0;
                docxStore.RegisterDocxTemplate<Doc>(template);
                docx.Add(Convert.ToHexString(await RenderBytes(docxStore)));

                var markdownStore = new TemplateStore();
                styleSource.Position = 0;
                markdownStore.RegisterMarkdownTemplate<Doc>(content, styleSource);
                markdown.Add(Convert.ToHexString(await RenderBytes(markdownStore)));
            }
        }

        await Assert.That(docx.Count).IsEqualTo(1);
        await Assert.That(markdown.Count).IsEqualTo(1);
    }

    // A Local value names an instant on the machine that built it, so what it prints follows that
    // machine's zone by definition. Pinned here so the documented behaviour does not drift.
    [Test]
    [NotInParallel]
    public async Task LocalKindFollowsTheMachineTimeZone()
    {
        using (LocalTimeZone.Emulate("Tokyo Standard Time"))
        {
            await Assert.That(await RenderDocx("{{ Local }}")).IsEqualTo("2026-03-01 04:05:09Z");
            await Assert.That(await RenderMarkdown("{{ Local }}")).IsEqualTo("2026-03-01 04:05:09Z");
            await Assert.That(await RenderDocx($"{{{{ Local | date: {stamp} }}}}")).IsEqualTo("2026-03-01 13:05:09 +0900");
            await Assert.That(await RenderMarkdown($"{{{{ Local | date: {stamp} }}}}")).IsEqualTo("2026-03-01 13:05:09 +0900");
        }
    }

    // "now" and "today" read the rendering machine's clock, zone included. Not deterministic by
    // nature; the readme says so.
    [Test]
    [NotInParallel]
    public async Task NowFollowsTheMachineClock()
    {
        using (LocalTimeZone.Emulate("Tokyo Standard Time"))
        {
            await Assert.That(await RenderDocx("{{ 'now' | date: '%z' }}")).IsEqualTo("+0900");
            await Assert.That(await RenderMarkdown("{{ 'today' | date: '%z' }}")).IsEqualTo("+0900");
        }
    }

    static async Task<string> RenderDocx(string content)
    {
        using var template = DocxTemplateBuilder.Build(content);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Doc>(template);
        return BodyText(await RenderBytes(store));
    }

    static async Task<string> RenderMarkdown(string content)
    {
        using var styleSource = DocxTemplateBuilder.Build();
        var store = new TemplateStore();
        store.RegisterMarkdownTemplate<Doc>(content, styleSource);
        return BodyText(await RenderBytes(store));
    }

    static async Task<byte[]> RenderBytes(TemplateStore store)
    {
        using var stream = new MemoryStream();
        await store.Render(new Doc(), stream);
        return stream.ToArray();
    }

    static string BodyText(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return string.Concat(document.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(_ => _.Text));
    }

    /// <summary>
    /// Swaps <see cref="TimeZoneInfo.Local"/> for the life of the scope. The runtime has no
    /// supported way to do this in-process (and no environment switch on Windows), so it reaches
    /// into the cache <see cref="TimeZoneInfo"/> keeps its local zone in. Only for tests marked
    /// <c>[NotInParallel]</c>.
    /// </summary>
    sealed class LocalTimeZone : IDisposable
    {
        static readonly object cachedData = typeof(TimeZoneInfo)
            .GetField("s_cachedData", BindingFlags.NonPublic | BindingFlags.Static)!
            .GetValue(null)!;

        static readonly FieldInfo localField = cachedData
            .GetType()
            .GetField("_localTimeZone", BindingFlags.NonPublic | BindingFlags.Instance)!;

        readonly TimeZoneInfo original = TimeZoneInfo.Local;

        public static LocalTimeZone Emulate(string id)
        {
            var scope = new LocalTimeZone();
            localField.SetValue(cachedData, TimeZoneInfo.FindSystemTimeZoneById(id));
            return scope;
        }

        public void Dispose() =>
            localField.SetValue(cachedData, original);
    }
}
