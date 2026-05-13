class RegisteredDocxTemplate(
    string name,
    Type modelType,
    byte[] canonicalBytes,
    IReadOnlyList<PartScopeTree> parts,
    ExcelsiorTableMap excelsiorTables,
    FormatMap formats,
    StringListMap stringLists,
    ImagePolicies imagePolicies) :
    RegisteredTemplate(name, modelType)
{
    public override async Task Render(object model, Stream output, Cancel cancel)
    {
        cancel.ThrowIfCancellationRequested();

        var context = new TemplateContext(model, SharedFluid.Options, allowModelMembers: true);
        using var stream = DocxCloner.ToWritableStream(canonicalBytes);
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var mainPart = doc.MainDocumentPart!;
            var numberingState = new WordNumberingState(mainPart);
            // Cache the StyleSet per render — Excelsior / Format / OpenXml / Mutate tokens all
            // need it, but it changes only across registrations, not within a render.
            var styles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));

            foreach (var part in parts)
            {
                cancel.ThrowIfCancellationRequested();
                await RenderPartAsync(doc, mainPart, part, context, model, numberingState, styles);
            }

            foreach (var (_, root) in DocxCloner.EnumerateParts(doc))
            {
                Anchors.StripAll(root);
            }

            doc.Save();
        }

        stream.Position = 0;
        await stream.CopyToAsync(output, cancel);
    }

    async Task RenderPartAsync(WordprocessingDocument doc, MainDocumentPart mainPart, PartScopeTree part, TemplateContext context, object model, WordNumberingState numberingState, Lazy<StyleSet> styles)
    {
        OpenXmlCompositeElement? root = null;
        foreach (var (uri, candidate) in DocxCloner.EnumerateParts(doc))
        {
            if (uri == part.PartUri)
            {
                root = candidate;
                break;
            }
        }

        if (root == null)
        {
            return;
        }

        var map = Anchors.BuildMap(root);
        var runner = new ScopeTreeRunner(
            Name,
            part.PartUri,
            map,
            context,
            mainPart,
            model,
            excelsiorTables,
            formats,
            stringLists,
            numberingState,
            styles,
            imagePolicies);
        await runner.RunAsync(part.Nodes);
        runner.ApplyStructural();
    }
}
