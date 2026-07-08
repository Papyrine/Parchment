using JetBrains.Annotations;

namespace Parchment;

/// <summary>
/// Declares that the decorated <c>partial</c> class is the binding model for a Parchment
/// template, validated at compile time by the Parchment source generator. The template path is
/// resolved against <c>AdditionalFiles</c>. The attribute is applied directly to the model class
/// — there is no separate marker / "template" class. See CLAUDE.md → "Design decisions".
///
/// The <c>[MeansImplicitUse]</c> annotation tells ReSharper / Rider that members of the
/// decorated class are bound implicitly at render time, so it stops emitting
/// "Property is never used" / <c>UnusedAutoPropertyAccessor.Global</c> warnings on them.
/// </summary>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
[MeansImplicitUse(ImplicitUseTargetFlags.Members)]
public sealed class ParchmentModelAttribute(string templatePath) :
    Attribute
{
    public string TemplatePath { get; } = templatePath;

    /// <summary>
    /// Controls document lockdown for templates whose model declares
    /// <see cref="EditableFieldAttribute"/> members. Passed through to
    /// <see cref="TemplateStore.RegisterDocxTemplate{TModel}(string, string, ProtectionMode)"/>
    /// by the generated <c>RegisterWith</c> helper. Ignored for markdown templates.
    /// </summary>
    public ProtectionMode Protection { get; set; } = ProtectionMode.WhenEditable;
}
