using System.Globalization;
// ReSharper disable PartialTypeWithSinglePart

public partial class EditableFieldTests
{
    [ParchmentBindable]
    public partial class EditableMeasures
    {
        [EditableField]
        public decimal Price { get; set; }

        [EditableField]
        public double Ratio { get; set; }

        [EditableField]
        public double Huge { get; set; }

        [EditableField]
        public float Weight { get; set; }

        [EditableField]
        public int Count { get; set; }

        [EditableField]
        public DateTimeOffset SignedAt { get; set; }

        [EditableField(DateFormat = "d MMMM yyyy HH:mm zzz")]
        public DateTimeOffset CountersignedAt { get; set; }

        [EditableField]
        public Time Pickup { get; set; }

        [EditableField(DateFormat = "h:mm tt")]
        public Time Dropoff { get; set; }

        [EditableField(DateFormat = "d MMMM yyyy")]
        public Date Due { get; set; }
    }

    const string measuresTemplate =
        """
        {{ Price }}

        {{ Ratio }}

        {{ Huge }}

        {{ Weight }}

        {{ Count }}

        {{ SignedAt }}

        {{ CountersignedAt }}

        {{ Pickup }}

        {{ Dropoff }}

        {{ Due }}
        """;

    static EditableMeasures NewMeasures() =>
        new()
        {
            Price = 12.50m,
            Ratio = 0.1,
            Huge = 1e32,
            Weight = 1234.5f,
            Count = 1000,
            SignedAt = new(2026, 7, 6, 9, 0, 0, TimeSpan.FromHours(10)),
            CountersignedAt = new(2026, 8, 7, 15, 30, 0, TimeSpan.FromHours(-5)),
            Pickup = new(16, 45, 0),
            Dropoff = new(17, 5, 0),
            Due = new(2026, 3, 4)
        };

    // th-TH and ar-SA default to non-Gregorian calendars; fr-FR and de-DE read '.' as a group
    // separator and name months differently.
    static readonly string[] extractCultures = ["", "en-US", "de-DE", "fr-FR", "th-TH", "ar-SA"];

    [Test]
    public async Task UneditedDocumentExtractsRenderedValuesUnderAnyCulture()
    {
        var expected = NewMeasures();
        using var stream = await RenderModel(measuresTemplate, expected);

        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            // Without w:fullDate the date control is read from its display text.
            var body = doc.MainDocumentPart!.Document!.Body!;
            FindSdt(body, "Due").SdtProperties!.GetFirstChild<SdtContentDate>()!.FullDate = null;
            doc.Save();
        }

        foreach (var name in extractCultures)
        {
            stream.Position = 0;
            var result = ParchmentExtractor.Extract<EditableMeasures>(stream, CultureInfo.GetCultureInfo(name));
            var failed = result.Fields
                .Where(_ => _.State != FieldState.Extracted)
                .Select(_ => $"{name}: {_.Path} {_.State} '{_.RawText}'")
                .ToList();
            await Assert.That(failed).IsEmpty();

            var actual = new EditableMeasures();
            result.ApplyTo(actual);
            await Assert.That(actual.Price).IsEqualTo(expected.Price);
            await Assert.That(actual.Ratio).IsEqualTo(expected.Ratio);
            await Assert.That(actual.Huge).IsEqualTo(expected.Huge);
            await Assert.That(actual.Weight).IsEqualTo(expected.Weight);
            await Assert.That(actual.Count).IsEqualTo(expected.Count);
            await Assert.That(actual.SignedAt).IsEqualTo(expected.SignedAt);
            await Assert.That(actual.SignedAt.Offset).IsEqualTo(expected.SignedAt.Offset);
            await Assert.That(actual.CountersignedAt).IsEqualTo(expected.CountersignedAt);
            await Assert.That(actual.Pickup).IsEqualTo(expected.Pickup);
            await Assert.That(actual.Dropoff).IsEqualTo(expected.Dropoff);
            await Assert.That(actual.Due).IsEqualTo(expected.Due);
        }
    }

    [Test]
    public async Task TypedNumberIsReadWithoutGroupSeparators()
    {
        // (field, typed text, extract culture, expected value or null for ParseFailed)
        (string Path, string Typed, string Culture, object? Expected)[] cases =
        [
            // A comma is never a group separator: under a culture without a decimal comma it
            // fails instead of being dropped.
            ("Price", "12,5", "", null),
            ("Price", "1,234", "", null),
            ("Price", "1,234.56", "", null),
            ("Price", "1,2,3", "en-US", null),
            ("Ratio", "12,5", "", null),
            ("Weight", "12,5", "en-US", null),

            // The caller's culture adds its decimal separator.
            ("Price", "12,5", "de-DE", 12.5m),
            ("Price", "1,234", "de-DE", 1.234m),
            ("Ratio", "12,5", "fr-FR", 12.5),
            ("Weight", "12,5", "de-DE", 12.5f),

            // A '.' is the decimal point the render wrote, whatever the culture.
            ("Price", "12.50", "de-DE", 12.50m),
            ("Ratio", "0.1", "de-DE", 0.1),
            ("Weight", "3.0", "fr-FR", 3f),

            // Both separators at once is a grouped number in some culture: never accepted.
            ("Price", "1.234,56", "de-DE", null),
            ("Ratio", "1.234,56", "de-DE", null),
            ("Price", "1 234", "fr-FR", null),

            // Integers take digits and a sign only.
            ("Count", "1,000", "", null),
            ("Count", "1,000", "en-US", null),
            ("Count", "1.000", "de-DE", null),
            ("Count", "-5", "de-DE", -5),

            // Symbols are not numbers.
            ("Price", "$5", "en-US", null),
            ("Price", "5 €", "de-DE", null),
            ("Ratio", "50%", "", null),
            ("Price", "-5", "", -5m),
            ("Ratio", "1e3", "", 1000d)
        ];

        using var stream = await RenderModel(measuresTemplate, NewMeasures());
        foreach (var (path, typed, culture, expected) in cases)
        {
            stream.Position = 0;
            using (var doc = WordprocessingDocument.Open(stream, true))
            {
                SetSdtText(FindSdt(doc.MainDocumentPart!.Document!.Body!, path), typed);
                doc.Save();
            }

            stream.Position = 0;
            var result = ParchmentExtractor.Extract<EditableMeasures>(stream, CultureInfo.GetCultureInfo(culture));
            var field = result.Fields.Single(_ => _.Path == path);
            var actual = $"{path} '{typed}' [{culture}] -> {field.State} {Describe(field.Value)}";
            var state = FieldState.Extracted;
            if (expected == null)
            {
                state = FieldState.ParseFailed;
            }

            await Assert.That(actual).IsEqualTo($"{path} '{typed}' [{culture}] -> {state} {Describe(expected)}");
            await Assert.That(field.RawText).IsEqualTo(typed);
        }
    }

    static string Describe(object? value)
    {
        if (value is IFormattable formattable)
        {
            return $"{value.GetType().Name} {formattable.ToString(null, CultureInfo.InvariantCulture)}";
        }

        return "-";
    }
}
