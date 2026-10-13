// ReSharper disable PartialTypeWithSinglePart

// An editable date is a wall-clock value: none of these depend on the machine's time zone. The
// zone is always explicit, either as the Kind of the rendered value or as the suffix written
// into w:fullDate.
public partial class EditableDateKindTests
{
    public class Slot
    {
        [EditableField]
        public DateTime At { get; set; }

        [EditableField(DateFormat = "yyyy-MM-dd HH:mm")]
        public DateTime Until { get; set; }

        [EditableField(DateFormat = "yyyy-MM-dd HH:mm")]
        public DateTimeOffset Logged { get; set; }
    }

    [ParchmentBindable]
    public partial class Schedule
    {
        [EditableField]
        public DateTime Starts { get; set; }

        [EditableField(DateFormat = "yyyy-MM-dd HH:mm")]
        public DateTime Ends { get; set; }

        [EditableField]
        public DateTime? Reviewed { get; set; }

        [EditableField]
        public Date Day { get; set; }

        [EditableField]
        public DateTimeOffset Signed { get; set; }

        [EditableField(DateFormat = "yyyy-MM-dd HH:mm")]
        public DateTimeOffset Logged { get; set; }

        [EditableField]
        public List<Slot> Slots { get; set; } = [];
    }

    const string template =
        """
        {{ Starts }} {{ Ends }} {{ Reviewed }} {{ Day }} {{ Signed }} {{ Logged }}

        {% for s in Slots %}

        {{ s.At }} {{ s.Until }} {{ s.Logged }}

        {% endfor %}
        """;

    static readonly DateTime midnight = new(2026, 7, 6);
    static readonly DateTime afternoon = new(2026, 7, 6, 14, 30, 15);
    static readonly DateTime morning = new(2026, 7, 6, 1, 2, 3);
    static readonly TimeSpan plusTen = TimeSpan.FromHours(10);

    static Schedule NewSchedule(DateTimeKind kind) =>
        new()
        {
            Starts = DateTime.SpecifyKind(midnight, kind),
            Ends = DateTime.SpecifyKind(afternoon, kind),
            Reviewed = DateTime.SpecifyKind(morning, kind),
            Day = new(2026, 7, 6),
            Signed = new(2026, 7, 6, 9, 0, 0, plusTen),
            Logged = new(2026, 7, 6, 9, 0, 0, plusTen),
            Slots =
            [
                new()
                {
                    At = DateTime.SpecifyKind(midnight, kind),
                    Until = DateTime.SpecifyKind(afternoon, kind),
                    Logged = new(2026, 7, 6, 9, 0, 0, plusTen)
                }
            ]
        };

    [Test]
    [Arguments(DateTimeKind.Unspecified)]
    [Arguments(DateTimeKind.Utc)]
    [Arguments(DateTimeKind.Local)]
    public async Task FullDateIsTheClockReadingWithNoZone(DateTimeKind kind)
    {
        using var stream = await Render(NewSchedule(kind));

        using var doc = WordprocessingDocument.Open(stream, false);
        var dates = DateControls(doc)
            .Select(_ => $"{Tag(_)}={FullDate(_).FullDate!.InnerText} [{_.InnerText}]")
            .ToList();

        await Assert.That(dates).IsEquivalentTo(
        [
            "Starts=2026-07-06T00:00:00 [2026-07-06]",
            "Ends=2026-07-06T14:30:15 [2026-07-06 14:30]",
            "Reviewed=2026-07-06T01:02:03 [2026-07-06]",
            "Day=2026-07-06T00:00:00 [2026-07-06]",
            "At=2026-07-06T00:00:00 [2026-07-06]",
            "Until=2026-07-06T14:30:15 [2026-07-06 14:30]"
        ]);
    }

    [Test]
    [Arguments(DateTimeKind.Unspecified)]
    [Arguments(DateTimeKind.Utc)]
    [Arguments(DateTimeKind.Local)]
    public async Task UntouchedDocumentRoundTripsTheClockReading(DateTimeKind kind)
    {
        using var stream = await Render(NewSchedule(kind));

        await AssertExtractsTheClockReading(stream);
    }

    // What a Word save leaves behind: "Z" on every w:fullDate (Word also rewrites an offset to the
    // clock reading plus "Z"). The offsets cover a document written by something other than Word,
    // or by an earlier Parchment for a Local value.
    [Test]
    [Arguments("Z")]
    [Arguments("+10:00")]
    [Arguments("-05:00")]
    public async Task ZoneSuffixOnFullDateDoesNotShiftTheValue(string suffix)
    {
        using var stream = await Render(NewSchedule(DateTimeKind.Local));

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            foreach (var control in DateControls(doc))
            {
                var fullDate = FullDate(control).FullDate!;
                fullDate.InnerText += suffix;
            }

            doc.Save();
        }

        stream.Position = 0;
        await AssertExtractsTheClockReading(stream);
    }

    [Test]
    public async Task DateTimeOffsetKeepsItsOffsetWhenTheTextCarriesOne()
    {
        using var stream = await Render(NewSchedule(DateTimeKind.Unspecified));

        var model = Extract(stream);

        await Assert.That(model.Signed.DateTime).IsEqualTo(new(2026, 7, 6, 9, 0, 0));
        await Assert.That(model.Signed.Offset).IsEqualTo(plusTen);
    }

    [Test]
    public async Task DateTimeOffsetWithoutAnOffsetInItsFormatExtractsAtOffsetZero()
    {
        using var stream = await Render(NewSchedule(DateTimeKind.Unspecified));

        var model = Extract(stream);

        // "yyyy-MM-dd HH:mm" shows no offset, so the clock reading is what survives; it comes
        // back at offset zero rather than at the offset of the machine doing the extraction.
        await Assert.That(model.Logged.DateTime).IsEqualTo(new(2026, 7, 6, 9, 0, 0));
        await Assert.That(model.Logged.Offset).IsEqualTo(TimeSpan.Zero);
        await Assert.That(model.Slots[0].Logged.DateTime).IsEqualTo(new(2026, 7, 6, 9, 0, 0));
        await Assert.That(model.Slots[0].Logged.Offset).IsEqualTo(TimeSpan.Zero);
    }

    static async Task AssertExtractsTheClockReading(Stream stream)
    {
        var model = Extract(stream);
        var slot = model.Slots.Single();

        await Assert.That(model.Starts.Ticks).IsEqualTo(midnight.Ticks);
        await Assert.That(model.Starts.Kind).IsEqualTo(DateTimeKind.Unspecified);
        await Assert.That(model.Ends.Ticks).IsEqualTo(afternoon.Ticks);
        await Assert.That(model.Ends.Kind).IsEqualTo(DateTimeKind.Unspecified);
        await Assert.That(model.Reviewed!.Value.Ticks).IsEqualTo(morning.Ticks);
        await Assert.That(model.Reviewed!.Value.Kind).IsEqualTo(DateTimeKind.Unspecified);
        await Assert.That(model.Day).IsEqualTo(new(2026, 7, 6));
        await Assert.That(slot.At.Ticks).IsEqualTo(midnight.Ticks);
        await Assert.That(slot.At.Kind).IsEqualTo(DateTimeKind.Unspecified);
        await Assert.That(slot.Until.Ticks).IsEqualTo(afternoon.Ticks);
        await Assert.That(slot.Until.Kind).IsEqualTo(DateTimeKind.Unspecified);
    }

    static Schedule Extract(Stream stream)
    {
        var result = ParchmentExtractor.Extract<Schedule>(stream);
        var model = new Schedule();
        result.ApplyTo(model);
        return model;
    }

    static List<SdtElement> DateControls(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!
            .Descendants<SdtElement>()
            .Where(_ => _.SdtProperties?.GetFirstChild<SdtContentDate>() != null)
            .ToList();

    static SdtContentDate FullDate(SdtElement sdt) =>
        sdt.SdtProperties!.GetFirstChild<SdtContentDate>()!;

    static string? Tag(SdtElement sdt) =>
        sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    static async Task<MemoryStream> Render(Schedule model)
    {
        using var docx = DocxTemplateBuilder.Build(template);
        var store = new TemplateStore();
        store.RegisterDocxTemplate<Schedule>(docx);

        var stream = new MemoryStream();
        await store.Render(model, stream);
        stream.Position = 0;
        return stream;
    }
}
