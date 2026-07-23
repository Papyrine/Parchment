// The models live in ParchmentSample, where they are decorated with [ParchmentModel] and their
// templates are registered as AdditionalFiles. That project is what compiles the readme's
// source-generator examples, so they cannot drift from the generator's actual output.
public class GeneratedRegistrationTests
{
    [Test]
    public async Task RegisterWithCoversBothFlows()
    {
        #region GeneratorRegisterWith

        var store = new TemplateStore();
        Invoice.RegisterWith(store);
        ReportContext.RegisterWith(store);

        #endregion

        using var invoice = new MemoryStream();
        await store.Render("Invoice", SampleData.Invoice(), invoice);
        invoice.Position = 0;

        using var report = new MemoryStream();
        await store.Render("ReportContext", SampleData.Report(), report);
        report.Position = 0;

        await Assert.That(Text(invoice)).Contains("Invoice INV-2026-0042");
        await Assert.That(Text(report)).Contains("Q2 Platform Health Review");
    }

    static string Text(Stream stream)
    {
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document!.Body!.InnerText;
    }
}
