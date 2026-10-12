namespace Parchment.Generated;

/// <summary>
/// Carries the bytes the source generator embedded for a model — a docx template, or a markdown
/// template's style document — as base64, on the type the generator nests in the model.
/// </summary>
/// <remarks>
/// An attribute argument rather than a string the generated code loads, because of where the
/// compiler stores each: a loaded string goes to the metadata user-string heap, as UTF-16, which
/// every literal of the assembly shares and which is limited to 16 MB; an attribute argument goes
/// to the blob heap, as UTF-8, which is not. Nothing reads it until a
/// <see cref="TemplateStore"/> first renders the model.
///
/// Not intended for hand-written consumption — it is emitted by the Parchment source generator.
/// </remarks>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
public sealed class EmbeddedBytesAttribute(string base64) :
    Attribute
{
    public string Base64 { get; } = base64;
}
