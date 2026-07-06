class RegisteredDocxTemplate(
    string name,
    Type modelType,
    byte[] canonicalBytes,
    IReadOnlyList<PartScopeTree> parts,
    ExcelsiorTableMap excelsiorTables,
    FormatMap formats,
    StringListMap stringLists,
    EditableMap editables,
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
            var editableState = new EditableState();
            // Cache the StyleSet per render — Excelsior / Format / OpenXml / Mutate tokens all
            // need it, but it changes only across registrations, not within a render.
            var styles = new Lazy<StyleSet>(() => StyleSet.Read(mainPart));

            foreach (var part in parts)
            {
                cancel.ThrowIfCancellationRequested();
                await RenderPartAsync(doc, mainPart, part, context, model, numberingState, editableState, styles);
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

    async Task RenderPartAsync(WordprocessingDocument doc, MainDocumentPart mainPart, PartScopeTree part, TemplateContext context, object model, WordNumberingState numberingState, EditableState editableState, Lazy<StyleSet> styles)
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

        // Editable fields dispatch only in the document body. Word does not reliably honor
        // editable-range exceptions in headers/footers/notes, so the same token there renders
        // as plain read-only text — deliberate (e.g. an editable PO number in the body can be
        // mirrored read-only in the footer).
        var isBody = part.PartUri == mainPart.Uri.ToString();

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
            isBody ? editables : EditableMap.Empty,
            editableState,
            numberingState,
            styles,
            imagePolicies);
        await runner.RunAsync(part.Nodes);
        runner.ApplyStructural();
    }
}
