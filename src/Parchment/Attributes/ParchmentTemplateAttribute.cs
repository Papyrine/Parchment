namespace Parchment;

/// <summary>
/// Declares that a partial class is backed by a docx template validated at compile time
/// by the Parchment source generator. The template path is resolved against AdditionalFiles.
/// </summary>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
public sealed class ParchmentTemplateAttribute(string templatePath, Type modelType) :
    Attribute
{
    public string TemplatePath { get; } = templatePath;
    public Type ModelType { get; } = modelType;
}
