// The models live in ParchmentSample, where they are decorated with [ParchmentModel] and their
// templates are registered as AdditionalFiles. That project is what compiles the readme's
// source-generator examples, so they cannot drift from the generator's actual output.
public class GeneratedRegistrationTests
{
    [Test]
    public async Task GeneratedModuleInitializerCoversBothFlows()
    {
        #region GeneratorRender

        // No registration call: the generated module initializer stored each model's embedded
        // template when the model assembly loaded, and the store picks it up on first render.
        var store = new TemplateStore();

        using var invoice = new MemoryStream();
        await store.Render(SampleData.Invoice(), invoice);

        #endregion

        invoice.Position = 0;

        using var report = new MemoryStream();
        await store.Render(SampleData.Report(), report);
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
