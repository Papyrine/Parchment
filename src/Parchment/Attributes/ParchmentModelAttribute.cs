using JetBrains.Annotations;

namespace Parchment;

/// <summary>
/// Declares that the decorated <c>partial</c> class is the binding model for a Parchment
/// template, validated at compile time by the Parchment source generator. The template is found
/// by convention: an <c>AdditionalFiles</c> entry named after the type — <c>TypeName.docx</c> or
/// <c>TypeName.md</c>. A markdown template's style source is <c>TypeName.dotx</c> when present,
/// otherwise the nearest <c>parchment.dotx</c> up the directory tree. The attribute is applied
/// directly to the model class — there is no separate marker / "template" class. See CLAUDE.md →
/// "Design decisions".
///
/// The <c>[MeansImplicitUse]</c> annotation tells ReSharper / Rider that members of the
/// decorated class are bound implicitly at render time, so it stops emitting
/// "Property is never used" / <c>UnusedAutoPropertyAccessor.Global</c> warnings on them.
/// </summary>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
[MeansImplicitUse(ImplicitUseTargetFlags.Members)]
public sealed class ParchmentModelAttribute :
    Attribute
{
    /// <summary>
    /// Controls document lockdown for templates whose model declares
    /// <see cref="EditableFieldAttribute"/> members. Passed through to the generated module
    /// initializer's registration. Ignored for markdown templates.
    /// </summary>
    public ProtectionMode Protection { get; set; } = ProtectionMode.WhenEditable;
}
